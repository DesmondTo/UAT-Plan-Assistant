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
