import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";
import { Label } from "@fluentui/react";
import { useId } from "@fluentui/react-hooks";

import FormActionContainer from "./FormActionContainer";

import { addActivity } from "../../utils/activityUtils/activityCreator";

function AddActivityForm() {
  const buttonId = useId("addActivityButton");
  const [activityTitle, setActivityTitle] = React.useState("");

  return (
    <form>
      <TextField
        label="Step 1: Enter your activity title:"
        value={activityTitle}
        onChange={(e) => setActivityTitle(e.target.value)}
      />
      <Label htmlFor={buttonId}>
        Step 2: Click on the cell you want to add your activity, then click the button below to add.
      </Label>
      <FormActionContainer>
        <DefaultButton
          id={buttonId}
          onClick={async () => {
            await addActivity(activityTitle);
          }}
        >
          Add Activity
        </DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default AddActivityForm;
