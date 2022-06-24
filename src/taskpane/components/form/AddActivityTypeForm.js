import React, { useState } from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";
import { Label } from "@fluentui/react";
import { useId } from "@fluentui/react-hooks";

import ProjectActivityDropdown from "../ProjectActivityDropdown";
import FormActionContainer from "./FormActionContainer";

import { addActivityType } from "../../utils/activityUtils/activityTypeCreator";

function AddActivityTypeForm() {
  const buttonId = useId("addActivityTypeButton");
  const [projectActivity, setProjectActivity] = useState();
  const [activityTypeTitle, setActivityTypeTitle] = useState("");

  return (
    <form>
      <ProjectActivityDropdown selectProjectActivity={setProjectActivity} />
      {projectActivity && (
        <>
          <TextField
            label="Step 2: Enter your activity type title:"
            value={activityTypeTitle}
            onChange={(e) => setActivityTypeTitle(e.target.value)}
          />
          <Label htmlFor={buttonId}>
            Step 3: Click on the cell you want to add your activity type, then click the button below to add.
          </Label>
          <FormActionContainer>
            <DefaultButton
              id={buttonId}
              onClick={async () => {
                await addActivityType(activityTypeTitle, projectActivity.address);
              }}
            >
              Add Activity Type
            </DefaultButton>
          </FormActionContainer>
        </>
      )}
    </form>
  );
}

export default AddActivityTypeForm;
