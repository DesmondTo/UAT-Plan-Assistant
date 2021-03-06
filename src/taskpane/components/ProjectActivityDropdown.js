import React, { useState, useEffect } from "react";

import { Stack } from "@fluentui/react/lib/Stack";
import { Dropdown } from "@fluentui/react/lib/Dropdown";

import { getAllProjectActivity } from "../utils/activityUtils/projectActivityGetter";

const dropdownStyles = {
  dropdown: { width: 300 },
};

const stackTokens = { childrenGap: 20 };

function ProjectActivityDropdown({ selectProjectActivity }) {
  const selectedProjectActivity = useState()[0];

  const onChange = (event, projectActivityObj) => {
    selectProjectActivity(projectActivityObj);
  };

  const [projectActivities, setProjectActivities] = useState([]);
  useEffect(async () => {
    const projectActivityArray = await getAllProjectActivity();
    setProjectActivities([...projectActivityArray]);
  }, []);

  return (
    <Stack tokens={stackTokens}>
      <Dropdown
        label="Step 1: Select an existing project activity:"
        selectedKey={selectedProjectActivity ? selectedProjectActivity.key : undefined}
        onChange={onChange}
        placeholder="Select an option"
        options={projectActivities}
        styles={dropdownStyles}
      />
    </Stack>
  );
}

export default ProjectActivityDropdown;
