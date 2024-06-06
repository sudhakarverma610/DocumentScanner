import { Button } from "@fluentui/react-components";
import * as React from "react";
import {
  CheckboxChecked16Regular,
  DocumentBorderPrint20Regular,
  DocumentArrowUp20Regular,
} from "@fluentui/react-icons";
import { HeaderButtonTypes } from "../../models/IconsumerForm";
export const Header = (props: {
  onButtonClick: (buttontype: HeaderButtonTypes) => void;
  disabled: boolean;
  showDisk: boolean;
  scanningViewExist: boolean;
}) => {
  return (
    <div style={{ display: "flex", gap: "25px" }}>
      {!props.scanningViewExist && (
        <Button
          onClick={() =>
            props.onButtonClick(HeaderButtonTypes.InitiateScanning)
          }
          appearance="primary"
          icon={<DocumentBorderPrint20Regular />}
          disabled={props.disabled}
        >
          Initiate Scanning
        </Button>
      )}
      {!props.scanningViewExist && (
        <Button
          onClick={() =>
            props.onButtonClick(HeaderButtonTypes.loadDiskDocument)
          }
          appearance="primary"
          icon={<DocumentArrowUp20Regular />}
          disabled={props.disabled}
        >
          Load Disk Document
        </Button>
      )}
      {props.scanningViewExist && (
        <Button
          onClick={() =>
            props.onButtonClick(HeaderButtonTypes.ScanningCompleted)
          }
          appearance="primary"
          icon={<CheckboxChecked16Regular />}
        >
          Scanning Completed
        </Button>
      )}
    </div>
  );
};
