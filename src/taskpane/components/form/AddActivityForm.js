import React, { useState } from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import ProjectActivityDropdown from "../ProjectActivityDropdown";
import FormActionContainer from "./FormActionContainer";

import { addActivity } from "../../utils/activityUtils/activityCreator";

function AddActivityForm() {
  const [projectActivity, setProjectActivity] = useState();
  const [activityTitle, setActivityTitle] = useState("");

  return (
    <form>
      <ProjectActivityDropdown selectProjectActivity={setProjectActivity} />
      {projectActivity && (
        <>
          <TextField
            label="Step 2: Enter your activity title:"
            value={activityTitle}
            onChange={(e) => setActivityTitle(e.target.value)}
          />
          <FormActionContainer>
            <DefaultButton
              onClick={async () => {
                await addActivity(activityTitle, projectActivity.address);
              }}
            >
              Add Activity
            </DefaultButton>
          </FormActionContainer>
        </>
      )}
    </form>
  );
}

export default AddActivityForm;
