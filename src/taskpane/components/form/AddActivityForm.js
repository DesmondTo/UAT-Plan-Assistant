import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import FormActionContainer from "./FormActionContainer";

function AddActivityForm() {
  const [activityTitle, setActivityTitle] = React.useState("");

  return (
    <form>
      <TextField
        label="Enter your activity title:"
        value={activityTitle}
        onChange={(e) => setActivityTitle(e.target.value)}
      />
      <FormActionContainer>
        <DefaultButton onClick={() => {}}>Add Activity</DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default AddActivityForm;
