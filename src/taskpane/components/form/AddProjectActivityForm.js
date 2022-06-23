import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import { addProjectActivity } from "../../utils/activityUtils/projectActivityCreator";
import FormActionContainer from "./FormActionContainer";

function AddProjectActivityForm() {
  const [projectActivityTitle, setProjectActivityTitle] = React.useState("");

  return (
    <form>
      <TextField
        label="Enter your project activity title:"
        value={projectActivityTitle}
        onChange={(e) => setProjectActivityTitle(e.target.value)}
      />
      <FormActionContainer>
        <DefaultButton onClick={() => addProjectActivity(projectActivityTitle)}>Add Project Activity</DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default AddProjectActivityForm;
