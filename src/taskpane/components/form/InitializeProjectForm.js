import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import { initializeProject } from "../../utils/projectUtils/projectInitializer";
import FormActionContainer from "./FormActionContainer";
/* global document, Office, module, require */

function InitializeProjectForm() {
  const [projectName, setProjectName] = React.useState("");
  const [kickOffDate, setKickOffDate] = React.useState();

  return (
    <>
      <form>
        <TextField
          label="Enter your project name:"
          value={projectName}
          onChange={(e) => setProjectName(e.target.value)}
        />
        <TextField
          type="date"
          label="Enter your project kick-off date:"
          value={kickOffDate}
          onChange={(e) => setKickOffDate(e.target.value)}
        />
        <FormActionContainer>
          <DefaultButton onClick={() => initializeProject(projectName, kickOffDate)}>Initialize Project</DefaultButton>
        </FormActionContainer>
      </form>
    </>
  );
}

export default InitializeProjectForm;
