/**
 * Returns date in {MMMM YYYY} format.
 * @param {string} dateString
 */
export const toShortDate = (dateString) => {
  const date = new Date(dateString);
  var month = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ][date.getMonth()];
  return month + " " + date.getFullYear();
};

/**
 * Returns date in {DD MMMM YYYY} format.
 * @param {string} dateString
 */
export const toLongDate = (dateString) => {
  return new Date(dateString).getDate() + " " + toShortDate(dateString);
};

/**
 * Returns weekday in 2 characters string format.
 * @param {number} weekdayNum Weekday number (0 - 7).
 */
export const toWeekDay = (weekdayNum) => {
  var dayOfWeek = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];
  return dayOfWeek[weekdayNum];
};

/**
 * Converts month string to number.
 * Zero indexed.
 * @param {string} monthStr 
 * @returns Month in number.
 */
export const toMonth = (monthStr) => {
  let monthNum = {
    January: 0,
    February: 1,
    March: 2,
    April: 3,
    May: 4,
    June: 5,
    July: 6,
    August: 7,
    September: 8,
    October: 9,
    November: 10,
    December: 11,
  };

  return monthNum[monthStr];
};
