import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import FormActionContainer from "./FormActionContainer";

function AddActivityTypeForm() {
  const [activityTypeTitle, setActivityTypeTitle] = React.useState("");

  return (
    <form>
      <TextField
        label="Enter your activity type title:"
        value={activityTypeTitle}
        onChange={(e) => setActivityTypeTitle(e.target.value)}
      />
      <FormActionContainer>
        <DefaultButton onClick={() => {}}>Add Activity Type</DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default AddActivityTypeForm;
