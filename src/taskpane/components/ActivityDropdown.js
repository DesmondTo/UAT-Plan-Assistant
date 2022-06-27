import React, { useState, useEffect } from "react";

import { Stack } from "@fluentui/react/lib/Stack";
import { Dropdown } from "@fluentui/react/lib/Dropdown";

import { getActivityOfProjectActivity } from "../utils/activityUtils/ActivityGetter";

const dropdownStyles = {
  dropdown: { width: 300 },
};

const stackTokens = { childrenGap: 20 };

function ActivityDropdown({ selectedProjectActivity, selectActivity }) {
  const selectedActivity = useState()[0];

  const onChange = (event, activityObj) => {
    selectActivity(activityObj);
  };

  const [activities, setActivities] = useState([]);

  useEffect(async () => {
    const activityArray = await getActivityOfProjectActivity(selectedProjectActivity.address);
    setActivities([...activityArray]);
  }, []);

  return (
    <Stack tokens={stackTokens}>
      <Dropdown
        label="Step 2: Select an existing project activity:"
        selectedKey={selectedActivity ? selectedActivity.key : undefined}
        onChange={onChange}
        placeholder="Select an option"
        options={activities}
        styles={dropdownStyles}
      />
    </Stack>
  );
}

export default ActivityDropdown;
