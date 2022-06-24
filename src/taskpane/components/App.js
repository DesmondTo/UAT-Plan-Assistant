import React, { useState } from "react";
import PropTypes from "prop-types";

import ActionDropdown from "./ActionDropdown";
import Progress from "./Progress";
import FormContainer from "./form/FormContainer";
import InitializeProjectForm from "./form/InitializeProjectForm";
import AddProjectActivityForm from "./form/AddProjectActivityForm";
import AddActivityTypeForm from "./form/AddActivityTypeForm";
import AddActivityForm from "./form/AddActivityForm";
import AddTimelineForm from "./form/AddTimelineForm";

const actionComponent = {
  addProj: <InitializeProjectForm />,
  addProjAct: <AddProjectActivityForm />,
  addActType: <AddActivityTypeForm />,
  addAct: <AddActivityForm />,
  addTl: <AddTimelineForm />,
};

/* global console, Excel, require */
export default function App({ title, isOfficeInitialized }) {
  const [action, setAction] = useState();

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  // let Component = actionComponent[action];

  return (
    <div className="ms-welcome">
      <ActionDropdown selectAction={setAction} />
      <FormContainer form={actionComponent[action]} />
    </div>
  );
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
