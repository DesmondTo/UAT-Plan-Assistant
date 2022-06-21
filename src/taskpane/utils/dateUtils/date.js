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

export const toLongDate = (dateString) => {
  return new Date(dateString).getDate() + " " + toShortDate(dateString);
};

export const toWeekDay = (dateObj) => {
  return dateObj.toLocaleString("default", { weekday: "short" }).substring(0, 2);
};
