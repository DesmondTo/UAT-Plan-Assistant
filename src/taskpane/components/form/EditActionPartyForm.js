import React from "react";

import { DefaultButton } from "@fluentui/react";
import { TextField } from "@fluentui/react";

import FormActionContainer from "./FormActionContainer";

function EditActionPartyForm() {
  const [newActionParty, setNewActionParty] = React.useState("");

  return (
    <form>
      <TextField
        label="Enter your new action party name:"
        value={newActionParty}
        onChange={(e) => setNewActionParty(e.target.value)}
      />
      <FormActionContainer>
        <DefaultButton onClick={() => {}}>Update Action Party</DefaultButton>
      </FormActionContainer>
    </form>
  );
}

export default EditActionPartyForm;
