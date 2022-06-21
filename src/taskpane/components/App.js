import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";

import { toLongDate, toShortDate, toWeekDay } from "../utils/dateUtils/date";

/* global console, Excel, require */
export default function App({ title, isOfficeInitialized }) {
  const [projectName, setProjectName] = React.useState("");
  const [kickOffDate, setKickOffDate] = React.useState();

  const initializeProjectCalendar = async () => {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const kickOffDateObj = new Date(kickOffDate);
      const kickOffYear = kickOffDateObj.getFullYear();
      const kickOffMonth = kickOffDateObj.getMonth() + 1;
      const kickOffDay = kickOffDateObj.getDate();
      const dayNums = new Date(kickOffYear, kickOffMonth, 0).getDate() - kickOffDay;

      const initialMonthRange = currentWorksheet.getRange("D1").getColumnsAfter(dayNums + 1);
      initialMonthRange.merge();
      currentWorksheet.getRange("E1").format.fill.color = "#8FBFF3";
      initialMonthRange.values = toShortDate(kickOffDate);
      initialMonthRange.numberFormat = "mmmm yyyy";
      initialMonthRange.format.font.bold = true;
      initialMonthRange.format.horizontalAlignment = "Center";

      const initialWeekDayRange = currentWorksheet.getRange("D2").getColumnsAfter(dayNums + 1);
      initialWeekDayRange.load("columnCount");
      await context.sync();
      for (var col = 0; col < initialWeekDayRange.columnCount; col++) {
        const currColumn = initialWeekDayRange.getColumn(col);
        currColumn.load(["format", "values"]);
        await context.sync();

        currColumn.values = toWeekDay(new Date(kickOffYear, kickOffMonth, kickOffDay + col));
        if (currColumn.values == "Sa") {
          currColumn.format.fill.color = "#E6B8AF";
        }
        if (currColumn.values == "Su") {
          currColumn.format.fill.color = "#DD7E6B";
        }
      }

      const initialDayRange = currentWorksheet.getRange("D3").getColumnsAfter(dayNums + 1);
      initialDayRange.load("columnCount");
      await context.sync();
      for (var col = 0; col < initialDayRange.columnCount; col++) {
        const currColumn = initialDayRange.getColumn(col);
        currColumn.load("values");
        await context.sync();

        currColumn.values = kickOffDay + col;
      }

      initialDayRange.format.autofitColumns();
      initialDayRange.format.autofitRows();
    });
  };

  const initializeProject = async () => {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.getRange().clear();
      const headerRange = currentWorksheet.getRange("A1:D4");
      const projectRange = currentWorksheet.getRange("A2:D1").load("columnCount");
      await context.sync();

      /* Project Name Header */
      headerRange.format.fill.color = "#1364BB";
      headerRange.format.font.bold = true;
      headerRange.format.font.color = "white";
      currentWorksheet.getRange("B2").format.horizontalAlignment = "Center";
      currentWorksheet.getRange("B2").values = `Project: ${projectName}`;

      /* Project Kick-off Date Header */
      currentWorksheet.getRange("B3:D4").format.fill.color = "white";
      currentWorksheet.getRange("B3:D4").format.font.color = "black";
      currentWorksheet.getRange("B3").values = `Kick-Off Date: ${toLongDate(kickOffDate)}`;

      /* Column Headers */
      currentWorksheet.getRange("C3").values = "Status";
      currentWorksheet.getRange("D3").values = "Action Party";

      currentWorksheet.getRange().format.autofitColumns();
      currentWorksheet.getRange().format.autofitRows();

      currentWorksheet.freezePanes.freezeAt("A1:D3");

      /* Update Project Columns Width */
      currentWorksheet.getCell(0, 0).format.columnWidth = 20;
      for (var col = 0; col < projectRange.columnCount; col++) {
        const currColumn = projectRange.getColumn(col);
        currColumn.load("format");
        await context.sync();

        currColumn.format.columnWidth = currColumn.format.columnWidth * 1.5;
      }

      /* Add Calendar Header */
      await initializeProjectCalendar();
      await context.sync();
    }).catch((error) => {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  };

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Hello" />
      <HeroList message="Start your project!" items={[]}>
        <form>
          <TextField
            label="Enter your project name:"
            value={projectName}
            onChange={(e) => setProjectName(e.target.value)}
          />
          <TextField
            type="date"
            label="Enter your project kick-off date:"
            value={kickOffDate}
            onChange={(e) => setKickOffDate(e.target.value)}
          />
        </form>
        <DefaultButton className="btn-danger" onClick={initializeProject}>
          Create Table
        </DefaultButton>
      </HeroList>
    </div>
  );
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
