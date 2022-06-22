import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import FormActionContainer from "./FormActionContainer";

function EditActivityTypeForm() {
  const [newActivityTypeTitle, setNewActivityTypeTitle] = React.useState("");

  return (
    <form>
      <TextField
        label="Enter your new activity type title:"
        value={newActivityTypeTitle}
        onChange={(e) => setNewActivityTypeTitle(e.target.value)}
      />
      <FormActionContainer>
        <DefaultButton onClick={() => {}}>Update Activity Type</DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default EditActivityTypeForm;
