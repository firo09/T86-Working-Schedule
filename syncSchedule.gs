/*******************************************************
 * 主函数：从表A(排班表)读取 => 全量写到表B的Sheet2 =>
 *       与Sheet1比对 => 若Sheet1无该日期则新建(生成eventID)，
 *       若有但班次不同则更新；最后清空Sheet2。
 *       
 * 根据eventID判断是否需要修改/创建日历事件，
 * 并在函数内部修正了 dateStr 可能是 Date 对象的问题。
 *******************************************************/
function syncYournameScheduleTwoTables() {
  // === 0) 可选：日历ID（若不需要同步到日历，可注释掉相关逻辑） ===
  const CALENDAR_ID = 'yourcalendarID@gmail.com';  // 替换为您自己的日历ID
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) {
    Logger.log(`警告：无法获取日历 ${CALENDAR_ID}。仍会执行Sheet对比，但无法写日历事件。`);
  }

  // === 1) 打开【排班表】，
  //      获取"T86 Working Schedule"并解析出 scheduleMap
  const SCHEDULE_SPREADSHEET_ID = 'yourSCHEDULE_SPREADSHEET_ID';
  const scheduleSS = SpreadsheetApp.openById(SCHEDULE_SPREADSHEET_ID);
  const scheduleSheet = scheduleSS.getSheetByName('T86 Working Schedule');
  if (!scheduleSheet) {
    throw new Error('找不到 T86 Working Schedule');
  }
  // 解析得到 { "MM/DD/YYYY": "OFF"/"Morning"/"Night"/"Work", ... }
  const scheduleMap = parseYournameSchedule(scheduleSheet);

  // === 2) 打开【映射/处理表】
  const MAPPING_SPREADSHEET_ID = 'yourMAPPING_SPREADSHEET_ID';
  const mappingSS = SpreadsheetApp.openById(MAPPING_SPREADSHEET_ID);

  // 2.1) 将scheduleMap“全量”写入Sheet2（不含eventID）
  let sheet2 = mappingSS.getSheetByName('Sheet2');
  if (!sheet2) {
    sheet2 = mappingSS.insertSheet('Sheet2');
    // 您可自行决定是否添加表头
    sheet2.appendRow(['Date', 'Shift Description']);
  }
  // 先清空Sheet2（除表头）再写入
  clearSheetButKeepHeader(sheet2);
  writeAllToSheet2(sheet2, scheduleMap);

  // 2.2) 读取/比对 Sheet2 与 Sheet1
  let sheet1 = mappingSS.getSheetByName('Sheet1');
  if (!sheet1) {
    // 若Sheet1不存在，则新建并添加标题行
    sheet1 = mappingSS.insertSheet('Sheet1');
    sheet1.appendRow(['Date', 'Shift Description', 'EventID']);
    Logger.log('已在映射表中新建Sheet1并添加表头');
  }

  // 从Sheet1读取 => { dateStr: { shiftDesc, eventId } }
  const sheet1Mapping = loadSheet1Mapping(sheet1);

  // 2.3) 遍历Sheet2里全部行 => 与Sheet1对比
  const lastRowSheet2 = sheet2.getLastRow();
  const dataSheet2 = sheet2.getDataRange().getValues(); // 包含表头
  let newCount = 0, updateCount = 0, skipCount = 0;

  for (let i = 1; i < lastRowSheet2; i++) {
    let [dateVal, shiftDesc] = dataSheet2[i];

    // 如果dateVal是日期对象，则转字符串 "MM/DD/YYYY"
    if (dateVal instanceof Date) {
      dateVal = formatDate(dateVal);
    }
    // 现在 dateVal 一定是字符串
    const dateStr = dateVal;

    // 判断Sheet1里有没有该日期
    const existing = sheet1Mapping[dateStr]; 
    if (!existing) {
      // ~~~ A) Sheet1中尚无该日期 => 新建并生成eventID
      let eventId = '';
      if (calendar) {
        // 新建日历事件
        const event = createEventForShift(calendar, dateStr, shiftDesc);
        eventId = event ? event.getId() : '';
      }
      // 在Sheet1中新建行
      appendToSheet1(sheet1, dateStr, shiftDesc, eventId);
      sheet1Mapping[dateStr] = { shiftDesc, eventId };
      newCount++;
    } else {
      // ~~~ B) 已存在 => 判断班次是否不同
      if (existing.shiftDesc !== shiftDesc) {
        // 班次变了 => 更新Sheet1
        updateSheet1Row(sheet1, dateStr, shiftDesc, existing.eventId);
        sheet1Mapping[dateStr].shiftDesc = shiftDesc;
        updateCount++;

        // 同时根据eventId是否存在 => 更新/重建对应日历事件
        if (calendar && existing.eventId) {
          const oldEvent = calendar.getEventById(existing.eventId);
          if (!oldEvent) {
            // 说明旧事件丢了 => 干脆重新创建
            const newEvent = createEventForShift(calendar, dateStr, shiftDesc);
            const newEid = newEvent ? newEvent.getId() : '';
            updateSheet1Row(sheet1, dateStr, shiftDesc, newEid);
            sheet1Mapping[dateStr].eventId = newEid;
            Logger.log(`日期${dateStr}的旧事件已丢失，重新创建了事件`);
          } else {
            // 如果 all-day<->timed 切换，需要删旧建新
            const shiftInfo = buildShiftInfo(shiftDesc);
            const updated = updateEventIfNeeded(oldEvent, shiftInfo);
            if (!updated) {
              // 删除旧事件
              oldEvent.deleteEvent();
              const newEvent = createEventForShift(calendar, dateStr, shiftDesc);
              const newEid = newEvent ? newEvent.getId() : '';
              updateSheet1Row(sheet1, dateStr, shiftDesc, newEid);
              sheet1Mapping[dateStr].eventId = newEid;
              Logger.log(`日期${dateStr}班次切换全日/时段 => 删旧建新事件`);
            } else {
              // 同类型 => 已更新标题/时段
              Logger.log(`日期${dateStr}的事件已更新`);
            }
          }
        }
      } else {
        // 班次不变 => 跳过
        skipCount++;
      }
    }
  }

  // 2.4) 对比结束 => 清空 Sheet2 (除了表头)
  clearSheetButKeepHeader(sheet2);

  // 日志输出
  Logger.log(`比对完成: 新增日期=${newCount}, 更新日期=${updateCount}, 跳过=${skipCount}`);
  Logger.log("已清空 Sheet2.");
}


/*********************************************************************
 *                            以下是辅助函数                          *
 *********************************************************************/

/**
 * 1) 从"T86 Working Schedule"中读取前200行，解析出 => { 'MM/DD/YYYY': shiftDesc, ... }
 *    只解析表头形如 "T86/AI SCHEDULE 12/25-1/5"
 *    检测到 "Yourname" 行后，把其单元格映射到当前日期段
 */
function parseYournameSchedule(scheduleSheet) {
  const allData = scheduleSheet.getDataRange().getValues().slice(0, 200);
  const scheduleMap = {};
  let currentDateList = [];

  for (let r = 0; r < allData.length; r++) {
    const row = allData[r];
    const firstCell = (row[0] || '').toString().trim();

    // (a) 检测表头: 形如 "T86/AI SCHEDULE 12/25-1/5"
    const headerPattern = /^T86\/AI SCHEDULE\s+(\d{1,2})\/(\d{1,2})\s*-\s*(\d{1,2})\/(\d{1,2})$/i;
    const match = headerPattern.exec(firstCell);
    if (match) {
      const sM = parseInt(match[1], 10);
      const sD = parseInt(match[2], 10);
      const eM = parseInt(match[3], 10);
      const eD = parseInt(match[4], 10);

      const [, startDate] = guessYearForMonthDay(sM, sD);
      const [, endDate]   = guessYearForMonthDay(eM, eD);

      const dateList = [];
      let d = new Date(startDate);
      while (d <= endDate) {
        dateList.push(new Date(d));
        d.setDate(d.getDate() + 1);
      }
      currentDateList = dateList;
      continue;
    }

    // (b) 若当前有日期范围，且第一列是 "Yourname"
    if (currentDateList.length > 0 && firstCell.toLowerCase() === 'yourname') {
      for (let i = 0; i < currentDateList.length; i++) {
        if (i + 1 >= row.length) break;
        const cellContent = (row[i+1] || '').toString().trim();
        const shiftInfo = parseShift(cellContent);

        // formatDate(...) => 得到 "MM/DD/YYYY"
        const dateStr = formatDate(currentDateList[i]); 
        scheduleMap[dateStr] = shiftInfo.title; // e.g. "OFF","Morning","Night","Work"
      }
    }
  }
  return scheduleMap;
}


/**
 * 2) 将 scheduleMap 全量写入 Sheet2，不含 eventID。
 *    假设 Sheet2 的列格式是 [Date, Shift Description]
 */
function writeAllToSheet2(sheet2, scheduleMap) {
  // 准备批量写入：[[dateStr, shiftDesc], ...]
  const rows = [];
  for (let dateStr in scheduleMap) {
    const shiftDesc = scheduleMap[dateStr];
    rows.push([dateStr, shiftDesc]);
  }
  if (rows.length > 0) {
    // 找到当前最后一行 + 1
    const startRow = sheet2.getLastRow() + 1;
    sheet2.getRange(startRow, 1, rows.length, 2).setValues(rows);
  }
  Logger.log(`已在Sheet2写入 ${rows.length} 条排班记录（日期->班次）。`);
}


/**
 * 3) 从Sheet1读取 => 生成 { dateStr: { shiftDesc, eventId }, ... }
 *    假设Sheet1的表头为 [Date, Shift Description, EventID]
 *    若读到的日期是 Date 对象，转成字符串 "MM/DD/YYYY"
 */
function loadSheet1Mapping(sheet1) {
  const data = sheet1.getDataRange().getValues();
  const dict = {};
  for (let i = 1; i < data.length; i++) {
    let [dateVal, shiftDesc, eventId] = data[i];
    if (!dateVal) continue;

    // 如果是 Date 对象 => 转字符串
    if (dateVal instanceof Date) {
      dateVal = formatDate(dateVal);
    }
    dict[dateVal] = {
      shiftDesc: shiftDesc || '',
      eventId: eventId || ''
    };
  }
  return dict;
}


/**
 * 4) 若Sheet1中没有 dateStr => 追加一行
 */
function appendToSheet1(sheet1, dateStr, shiftDesc, eventId) {
  // 如果 dateStr 可能是 Date对象，则先转字符串
  if (dateStr instanceof Date) {
    dateStr = formatDate(dateStr);
  }
  sheet1.appendRow([dateStr, shiftDesc, eventId]);
}

/**
 * 5) 更新已有行 => 修改Shift / eventId
 */
function updateSheet1Row(sheet1, dateStr, newShift, newEventId) {
  const data = sheet1.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    let [dt] = data[i];
    // 若 dt 是 Date => 转字符串
    if (dt instanceof Date) {
      dt = formatDate(dt);
    }
    if (dt === dateStr) {
      // 第 i+1 行, 第2列=>ShiftDesc, 第3列=>EventID
      sheet1.getRange(i+1, 2).setValue(newShift);
      if (newEventId !== undefined) {
        sheet1.getRange(i+1, 3).setValue(newEventId);
      }
      return;
    }
  }
}


/**
 * 清空Sheet(除表头), 仅保留第一行
 */
function clearSheetButKeepHeader(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}


/************************************************************************
 *                             日历事件相关                              *
 *        在此处统一处理 dateVal => dateStr，避免 .split 出错             *
 ************************************************************************/

/**
 * 创建(OFF=全天, 否则带时段)的事件，并返回该事件对象
 */
function createEventForShift(calendar, dateVal, shiftDesc) {
  if (!calendar) return null; 
  // 如果传进来是Date对象 => 转字符串
  if (dateVal instanceof Date) {
    dateVal = formatDate(dateVal);
  }

  const shiftInfo = buildShiftInfo(shiftDesc);
  // 确保 dateVal 是字符串 "MM/DD/YYYY"
  const [m, d, y] = dateVal.split('/').map(Number);

  if (shiftDesc === 'OFF') {
    const startDate = new Date(y, m-1, d);
    const ev = calendar.createAllDayEvent(shiftDesc, startDate, {
      description:`UniqueID:${dateVal}`
    });
    ev.setColor(getColorIdForShift(shiftDesc));
    return ev;
  } else {
    const startTime = new Date(y, m-1, d, shiftInfo.start.getHours(), shiftInfo.start.getMinutes());
    const endTime   = new Date(y, m-1, d, shiftInfo.end.getHours(),   shiftInfo.end.getMinutes());
    const ev = calendar.createEvent(shiftDesc, startTime, endTime, {
      description:`UniqueID:${dateVal}`
    });
    ev.setColor(getColorIdForShift(shiftDesc));
    ev.addPopupReminder(10);
    return ev;
  }
}

/**
 * 更新现有事件(若all-day <-> timed切换 => false表示需删旧建新)
 */
function updateEventIfNeeded(event, shiftInfo) {
  const oldAllDay = event.isAllDayEvent();
  const newAllDay = (shiftInfo.title === 'OFF');
  if (oldAllDay !== newAllDay) {
    return false; // 需要删旧建新
  }
  // 同类型 => 更新
  if (event.getTitle() !== shiftInfo.title) {
    event.setTitle(shiftInfo.title);
  }
  if (!newAllDay) {
    event.setTime(shiftInfo.start, shiftInfo.end);
  }
  event.setColor(getColorIdForShift(shiftInfo.title));
  event.removeAllReminders();
  if (!newAllDay) {
    event.addPopupReminder(10);
  }
  return true;
}

/**
 * 将 "OFF"/"Morning"/"Night"/"Work" => { title, start, end }
 */
function buildShiftInfo(shiftDesc) {
  switch (shiftDesc) {
    case 'OFF':
      return { title:'OFF', start:null, end:null };
    case 'Morning': {
      const s = new Date(); s.setHours(8,0,0,0);
      const e = new Date(); e.setHours(17,0,0,0);
      return { title:'Morning', start:s, end:e };
    }
    case 'Night': {
      const s = new Date(); s.setHours(13,0,0,0);
      const e = new Date(); e.setHours(22,0,0,0);
      return { title:'Night', start:s, end:e };
    }
    default: {
      // Work
      const s = new Date(); s.setHours(9,0,0,0);
      const e = new Date(); e.setHours(18,0,0,0);
      return { title:'Work', start:s, end:e };
    }
  }
}

/**
 * 为不同班次返回Google日历颜色ID
 */
function getColorIdForShift(shiftDesc) {
  switch (shiftDesc.toLowerCase()) {
    case 'morning': return '10'; // 绿色
    case 'night':   return '8';  // 灰色
    case 'off':     return '11'; // 红色
    case 'work':
    default:        return '1';  // 浅紫
  }
}


/************************************************************************
 *                             其它通用函数                              *
 ************************************************************************/

/**
 * (month,day) => 猜测哪一年(±3个月)
 */
function guessYearForMonthDay(m, d) {
  const now = new Date();
  const thisYear = now.getFullYear();
  const candidates = [thisYear-1, thisYear, thisYear+1];

  let best = thisYear;
  let bestDiff = Number.MAX_VALUE;
  for (let y of candidates) {
    const dt = new Date(y, m-1, d);
    const diff = Math.abs(dt - now);
    if (diff < bestDiff && diff <= 90*86400000) {
      best = y;
      bestDiff = diff;
    }
  }
  return [best, new Date(best, m-1, d)];
}

/**
 * 解析单元格内容 => { title, start, end }
 * - OFF => 全天
 * - Morning => 8-17
 * - Night => 13-22
 * - 其余(空白/work/both等) => 9-18
 */
function parseShift(cellContent) {
  const lower = (cellContent || '').toLowerCase();
  if (lower.includes('off'))      return { title: 'OFF',     start:null, end:null };
  if (lower.includes('morning'))  return { title: 'Morning', start:null, end:null };
  if (lower.includes('night'))    return { title: 'Night',   start:null, end:null };
  return { title: 'Work',         start:null, end:null };
}

/**
 * 格式化日期 => "MM/DD/YYYY"
 */
function formatDate(d) {
  const mm = ('0' + (d.getMonth()+1)).slice(-2);
  const dd = ('0' + d.getDate()).slice(-2);
  return `${mm}/${dd}/${d.getFullYear()}`;
}
