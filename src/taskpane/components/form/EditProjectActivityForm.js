import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import FormActionContainer from "./FormActionContainer";

function EditProjectActivityForm() {
  const [newProjectActivityTitle, setNewProjectActivityTitle] = React.useState("");

  return (
    <form>
      <TextField
        label="Enter your new project activity title:"
        value={newProjectActivityTitle}
        onChange={(e) => setNewProjectActivityTitle(e.target.value)}
      />
      <FormActionContainer>
        <DefaultButton onClick={() => {}}>Update Project Activity</DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default EditProjectActivityForm;
