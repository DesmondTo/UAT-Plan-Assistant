import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import FormActionContainer from "./FormActionContainer";

/* global document, Office, module, require */
function EditProjectForm() {
  const [newProjectName, setNewProjectName] = React.useState("");
  const [newKickOffDate, setNewKickOffDate] = React.useState();

  return (
    <form>
      <TextField
        label="Enter your new project name:"
        value={newProjectName}
        onChange={(e) => setNewProjectName(e.target.value)}
      />
      <TextField
        type="date"
        label="Enter your new project kick-off date:"
        value={newKickOffDate}
        onChange={(e) => setNewKickOffDate(e.target.value)}
      />
      <FormActionContainer>
        <DefaultButton onClick={() => initializeProject(newProjectName, newKickOffDate)}>Update Project</DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default EditProjectForm;
