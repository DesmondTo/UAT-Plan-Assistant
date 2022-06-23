import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";
import { Label } from "@fluentui/react";
import { useId } from "@fluentui/react-hooks";

import FormActionContainer from "./FormActionContainer";

import { addTimeline } from "../../utils/activityUtils/timelineCreator";

function AddTimelineForm() {
  const buttonId = useId("addTimelineButton");
  const [startDate, setStartDate] = React.useState();
  const [endDate, setEndDate] = React.useState();

  return (
    <form>
      <TextField
        type="date"
        label="Step 1: Enter your activity start date:"
        value={startDate}
        onChange={(e) => setStartDate(e.target.value)}
      />
      <TextField
        type="date"
        label="Step 2: Enter your activity end date:"
        value={endDate}
        onChange={(e) => setEndDate(e.target.value)}
      />
      <Label htmlFor={buttonId}>
        Step 3: Click on the activity you want to add to, then click the button below to add.
      </Label>
      <FormActionContainer>
        <DefaultButton
          id={buttonId}
          onClick={async () => {
            await addTimeline(startDate, endDate);
          }}
        >
          Add Timeline
        </DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default AddTimelineForm;
