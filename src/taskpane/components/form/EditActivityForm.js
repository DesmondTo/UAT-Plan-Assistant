import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import FormActionContainer from "./FormActionContainer";

function EditActivityForm() {
  const [newActivityTitle, setNewActivityTitle] = React.useState("");

  return (
    <form>
      <TextField
        label="Enter your new activity title:"
        value={newActivityTitle}
        onChange={(e) => setNewActivityTitle(e.target.value)}
      />
      <FormActionContainer>
        <DefaultButton onClick={() => {}}>Update Activity</DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default EditActivityForm;
