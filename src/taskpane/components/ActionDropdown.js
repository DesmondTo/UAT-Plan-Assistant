import React from "react";
import { Dropdown, DropdownMenuItemType } from "@fluentui/react/lib/Dropdown";

const dropdownStyles = { dropdown: { width: 300 }, padding: "20px" };

const dropdownOptions = [
  { key: "add", text: "Add", itemType: DropdownMenuItemType.Header },
  { key: "addProj", text: "Add Project" },
  { key: "addProjAct", text: "Add Project Activity" },
  { key: "addActType", text: "Add Activity Type" },
  { key: "addAct", text: "Add Activity" },
  { key: "addStat", text: "Add Status Key" },
  { key: "addActPar", text: "Add Action Party" },
  { key: "divider", text: "-", itemType: DropdownMenuItemType.Divider },
  { key: "edit", text: "Edit", itemType: DropdownMenuItemType.Header },
  { key: "editProj", text: "Edit Project" },
  { key: "editProjAct", text: "Edit Project Activity" },
  { key: "editActType", text: "Edit Activity Type" },
  { key: "editAct", text: "Edit Activity" },
  { key: "editStat", text: "Edit Status Key" },
  { key: "editActPar", text: "Edit Action Party" },
];

function ActionDropdown({ selectAction }) {
  const selectedItem = React.useState()[0];

  const onChange = (event, item) => {
    selectAction(item.key);
  };

  return (
    <Dropdown
      label="What to do?"
      selectedKey={selectedItem ? selectedItem.key : undefined}
      onChange={onChange}
      placeholder="Select an option"
      options={dropdownOptions}
      styles={dropdownStyles}
    />
  );
}

export default ActionDropdown;
