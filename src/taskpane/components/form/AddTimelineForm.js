import React, { useState } from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import ProjectActivityDropdown from "../ProjectActivityDropdown";
import ActivityDropdown from "../ActivityDropdown";
import FormActionContainer from "./FormActionContainer";

import { addTimeline } from "../../utils/activityUtils/timelineCreator";

function AddTimelineForm() {
  const [projectActivity, setProjectActivity] = useState();
  const [activity, setActivity] = useState();
  const [startDate, setStartDate] = useState();
  const [endDate, setEndDate] = useState();

  return (
    <form>
      <ProjectActivityDropdown selectProjectActivity={setProjectActivity} />
      {projectActivity && <ActivityDropdown selectedProjectActivity={projectActivity} selectActivity={setActivity} />}
      {projectActivity && activity && (
        <>
          <TextField
            type="date"
            label="Step 3: Enter your activity start date:"
            value={startDate}
            onChange={(e) => setStartDate(e.target.value)}
            required
          />
          <TextField
            type="date"
            label="Step 4: Enter your activity end date:"
            value={endDate}
            onChange={(e) => setEndDate(e.target.value)}
            required
          />
          <FormActionContainer>
            <DefaultButton
              onClick={async () => {
                await addTimeline(activity.address, startDate, endDate);
              }}
            >
              Add Timeline
            </DefaultButton>
          </FormActionContainer>
        </>
      )}
    </form>
  );
}

export default AddTimelineForm;
