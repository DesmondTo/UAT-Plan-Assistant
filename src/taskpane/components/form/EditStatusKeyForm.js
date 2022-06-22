import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import FormActionContainer from "./FormActionContainer";

function EditStatusKeyForm() {
  const [newStatusKey, setNewStatusKey] = React.useState("");
  const [newStatusColor, setNewStatusColor] = React.useState("");

  return (
    <form>
      <TextField
        label="Enter your status key title:"
        value={newStatusKey}
        onChange={(e) => setNewStatusKey(e.target.value)}
      />
      <TextField
        label="Enter your status key color:"
        value={newStatusColor}
        onChange={(e) => setNewStatusColor(e.target.value)}
      />
      <FormActionContainer>
        <DefaultButton onClick={() => {}}>Update Status Key</DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default EditStatusKeyForm;
