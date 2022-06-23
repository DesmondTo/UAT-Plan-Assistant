import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import ProjectActivityDropdown from "../ProjectActivityDropdown";
import FormActionContainer from "./FormActionContainer";

function AddActivityTypeForm() {
  const [projectActivity, setProjectActivity] = React.useState();
  const [activityTypeTitle, setActivityTypeTitle] = React.useState("");
  
  return (
    <form>
      <ProjectActivityDropdown selectProjectActivity={setProjectActivity} />
      {/* Render the form to add activity after project activity is selected */}
      {projectActivity && (
        <>
          <TextField
            label="Select which project activity to put your activity type:"
            value={activityTypeTitle}
            onChange={(e) => setActivityTypeTitle(e.target.value)}
          />
          <TextField
            label="Enter your activity type title:"
            value={activityTypeTitle}
            onChange={(e) => setActivityTypeTitle(e.target.value)}
          />
          <FormActionContainer>
            <DefaultButton onClick={() => {}}>Add Activity Type</DefaultButton>
          </FormActionContainer>
        </>
      )}
    </form>
  );
}

export default AddActivityTypeForm;
