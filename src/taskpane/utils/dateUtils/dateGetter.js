import { PROJECT_KICKOFF_DATE_CELL_INDEX } from "../../../constants/projectConstants";

import { toMonth } from "../dateUtils/dateFormatter";

/**
 * Gets the year of the specified date string.
 * @param {string} dateStr
 * @returns The year of the date, in number.
 */
export const getYear = (dateStr) => {
  return new Date(dateStr).getFullYear();
};

/**
 * Gets the month of the specified date string.
 * @param {string} dateStr
 * @returns The month of the date, in number.
 */
export const getMonth = (dateStr) => {
  return new Date(dateStr).getMonth();
};

/**
 * Gets the day of the month of the specified date string.
 * @param {string} dateStr
 * @returns The day of the date, in number.
 */
export const getDayOfTheMonth = (dateStr) => {
  return new Date(dateStr).getDate();
};

/**
 * Gets the number of days between two dates.
 * @param {date} date_1
 * @param {date} date_2
 * @returns The number of days between two dates.
 */
export const getDayDifference = (date_1, date_2) => {
  let difference = Math.abs(date_2.getTime() - date_1.getTime());
  let TotalDays = Math.ceil(difference / (1000 * 3600 * 24));
  return TotalDays;
};

/**
 * Adds a number of months to a date string.
 * @param {string} dateStr The string of the date, in the format of YYYY-MM-DD.
 * @param {number} monthNum The number of months to be addded to the date.
 * @returns The string of the date after increased by a number of months.
 */
export const addMonthToDateString = (dateStr, monthNum) => {
  const [year, month, day] = dateStr.split("-");
  const newMonth = +month + monthNum; // Converts the month to number, then add.
  return `${year}-${newMonth < 10 ? `0${newMonth}` : newMonth}-${day}`;
};

/**
 * Gets an array of date strings between two dates inclusively, each increased by one month.
 * @param {string} startDate
 * @param {string} endDate
 * @returns An array of date string between two dates.
 */
export const getDateStringArrayIncreasedByMonth = (startDate, endDate) => {
  let dateStringArr = [];
  let dateStringToAdd = startDate;
  while (getMonth(dateStringToAdd) !== getMonth(endDate)) {
    dateStringArr.push(dateStringToAdd);
    dateStringToAdd = addMonthToDateString(dateStringToAdd, 1);
  }
  dateStringArr.push(endDate);
  return dateStringArr;
};

/**
 * Gets the number of days between first day of kick-off month to specified date.
 * @param {string} dateStr
 * @returns Number of days between first day of kick-off month to specified date.
 */
export const getDayNumFromKickOffMonth = async (year, month) => {
  let dayNum = 0;

  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    let kickOffDate = currentWorksheet.getRange(PROJECT_KICKOFF_DATE_CELL_INDEX);
    kickOffDate.load("values");
    await context.sync();

    kickOffDate = kickOffDate.values[0][0].replace("Kick-Off Date: ", ""); // Get the date in the value array.
    const [kickOffDay, kickOffMonth, kickOffYear] = kickOffDate.split(" "); // Split the date into three entities.
    const kickOffDateObj = new Date(kickOffYear, toMonth(kickOffMonth), 1); // Month is in string, change it to number.
    dayNum = getDayDifference(kickOffDateObj, new Date(year, month, 1));
    await context.sync();
  });
  
  return dayNum;
};
