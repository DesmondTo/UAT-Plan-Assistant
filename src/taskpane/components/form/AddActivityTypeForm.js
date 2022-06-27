import React, { useState } from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import ProjectActivityDropdown from "../ProjectActivityDropdown";
import FormActionContainer from "./FormActionContainer";

import { addActivityType } from "../../utils/activityUtils/activityTypeCreator";

function AddActivityTypeForm() {
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
          <FormActionContainer>
            <DefaultButton
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
