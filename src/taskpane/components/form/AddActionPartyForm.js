import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import FormActionContainer from "./FormActionContainer";

function AddActionPartyForm() {
  const [actionParty, setActionParty] = React.useState("");

  return (
    <form>
      <TextField
        label="Enter your action party name:"
        value={actionParty}
        onChange={(e) => setActionParty(e.target.value)}
      />
      <FormActionContainer>
        <DefaultButton onClick={() => {}}>Add Action Party</DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default AddActionPartyForm;
