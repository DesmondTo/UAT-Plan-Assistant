import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import FormActionContainer from "./FormActionContainer";

function AddStatusKeyForm() {
  const [statusKey, setStatusKey] = React.useState("");
  const [statusColor, setStatusColor] = React.useState("");

  return (
    <form>
      <TextField
        label="Enter your status key title:"
        value={statusKey}
        onChange={(e) => setStatusKey(e.target.value)}
      />
      <TextField
        label="Enter your status key color:"
        value={statusColor}
        onChange={(e) => setStatusColor(e.target.value)}
      />
      <FormActionContainer>
        <DefaultButton onClick={() => {}}>Add Status Key</DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default AddStatusKeyForm;
