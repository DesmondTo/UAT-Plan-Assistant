import React from "react";
import PropTypes from "prop-types";

import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
import InitializeProjectForm from "./form/InitializeProjectForm";

/* global console, Excel, require */
export default function App({ title, isOfficeInitialized }) {
  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Hello" />
      <HeroList message="Start your project!" items={[]}>
        <InitializeProjectForm />
      </HeroList>
    </div>
  );
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
