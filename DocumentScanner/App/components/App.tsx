import { FC, ReactNode, useEffect, useState } from "react";
import { IInputs } from "../../generated/ManifestTypes";
import * as React from "react";
import Dynamsoft from "dwt";
import { ImageViewer } from "./ImageViewer/ImageViewer";
import {
  Button,
  FluentProvider,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  webLightTheme,
} from "@fluentui/react-components";
import {
  CheckboxChecked16Regular,
  DocumentBorderPrint20Regular,
  DocumentArrowUp20Regular,
} from "@fluentui/react-icons";
import "./App.css";
import { Header } from "./Header";
import { HeaderButtonTypes } from "../../models/IconsumerForm";

export interface AppProps {
  context: ComponentFramework.Context<IInputs>;
  appId: string;
  children?: ReactNode;
}
export const App: FC<AppProps> = (props: AppProps) => {
  const [DWObject, setDWObject] = useState<any>(null);
  const [scanningDone, setScanningDone] = useState<boolean>(false);
  const [isIRCDocumentRole, setisIRCDocumentRole] = useState<boolean>(false);

  const [ableToInitiateScanning, setAbleToInitiateScanning] = useState<boolean>(
    false
  );
  let containerId = "dwtcontrolContainer";
   
  const [showScanningView, setShowScanningView] = useState<boolean>(false); 
  const trailKey =
    "t01888AUAAGuTHmGf2O9qgopSfxULa5hbf3NTP1CGwiLkfYUQUeBjJuPuG31v6I059Qj5tOuVifpdctnaUaRcd5WTO1EHkml/Tp7glPFOofFOTHDyLSfRcxv6vNvmNy9O4Ai8NkD247ADqIFcSwO8bakNGsAaIAogWg3QgNNV9D+fqQ1I/fXKhiYnT3DKeGcdkDFOTHDyLWcIyOLJrGG1aw4I6pNzAFgD5BRAuci6gFALsAZIB8DCijfyBXHcKoU=";

  let Dynamsoft_OnReady = () => {
    setDWObject(Dynamsoft.DWT.GetWebTwain(containerId));
  };
  let onPostAllTransfers = () => {
    console.log("onPostAllTransfers");
  };
  let onLoadDWT = async () => {
    Dynamsoft.DWT.RegisterEvent("OnWebTwainReady", () => {
      Dynamsoft_OnReady();
      Dynamsoft.DWT.RegisterEvent("OnPostTransfer", onPostAllTransfers);
    });
    Dynamsoft.DWT.RegisterEvent("OnPreTransfer", function () {
      console.log("OnPreTransfer");
    });
    Dynamsoft.DWT.RegisterEvent("OnPostTransfer", function () {
      console.log("OnPostTransfer");
    });

    //
    Dynamsoft.DWT.ProductKey = await getConfig("Dynamsoft.DWT.ProductKey");
    Dynamsoft.DWT.ResourcesPath = "https://unpkg.com/dwt@18.4.2/dist"; //"WebResources/neu_/lib";
    Dynamsoft.DWT.Containers = [
      {
        WebTwainId: "dwtObject",
        ContainerId: containerId,
        Width: "500px",
        Height: "500px",
      },
    ];
    Dynamsoft.DWT.Load();
  };
  let acquireImage = () => {
    if (DWObject) {
      DWObject.SelectSourceAsync()
        .then(() => {
          return DWObject.AcquireImageAsync({
            IfCloseSourceAfterAcquire: true,
          });
        })
        .catch((exp: any) => {
          console.error(exp);
        })
        .then((x: any) => {
          console.log("scanning done", x);
          if (x) {
            setShowScanningView(true);
           }
        });
    }
  };
  useEffect(() => {
    var xrmUtility = (window as any).Xrm;
   let isIrc= xrmUtility.Utility.getGlobalContext()
    .userSettings.roles.getAll()
    .filter(
      (r: any) =>
        r.name.toLowerCase() == "IRC ADD Document Scanning".toLowerCase()       
    ).length > 0;
    if(isIrc){
      setisIRCDocumentRole(isIrc)
      onLoadDWT();
      isIntialButton();
    }
 
  }, []);
  let loadImage = () => {
    let OnSuccess = () => {
      setShowScanningView(true);
      console.log("successful");
    };

    let OnFailure = (errorCode: any, errorString: any) => {
      alert(errorString);
    };
    if (DWObject) {
      DWObject.IfShowFileDialog = true; // Open the system's file dialog to load image
      DWObject.LoadImageEx(
        "",
        Dynamsoft.DWT.EnumDWT_ImageType.IT_ALL,
        OnSuccess,
        OnFailure
      ); // Load images in all supported formats (.bmp, .jpg, .tif, .png, .pdf). OnSuccess or   OnFailure will be called after the operation
    }
  };
  let onDone = () => {
    console.log("onDone");
    setScanningDone(true);
    setShowScanningView(false);
  };
  let getConfig = async (name: string) => {
    try {
      var response = await props.context.webAPI.retrieveMultipleRecords(
        "neu_systemsetting",
        "?$filter=neu_name eq '" + name + "'&$select=neu_value"
      );
      return response.entities?.[0]?.neu_value;
    } catch (e) {
      return trailKey;
    }
  };
  let isIntialButton = async () => {
    var entityIdWithCurly = (props.context as any).page.entityId;
    if (!entityIdWithCurly) {
      setAbleToInitiateScanning(false);
      return;
    }
    var entityId = (props.context as any).page.entityId
      .replace("{", "")
      .replace("}", "");
    var userSettings = props.context.userSettings;
    var currentuserid = userSettings.userId.replace("{", "").replace("}", "");
    // setAbleToInitiateScanning(true);
    // return;
    var appendToAccess = await checkAccess("contact", entityId, currentuserid);
    var xrmUtility = (window as any).Xrm;
    if (appendToAccess && xrmUtility) {
      if (
        xrmUtility.Utility.getGlobalContext()
          .userSettings.roles.getAll()
          .filter(
            (r: any) =>
              r.name.toLowerCase() == "system administrator" ||
              r.name.toLowerCase() == "irc admin" ||
              r.name.toLowerCase() == "irc case control" ||
              r.name.toLowerCase() == "irc clerical support" ||
              r.name.toLowerCase() == "irc clinical" ||
              r.name.toLowerCase() == "irc director" ||
              r.name.toLowerCase() == "irc finance" ||
              r.name.toLowerCase() == "irc intake coordinator" ||
              r.name.toLowerCase() == "irc it" ||
              r.name.toLowerCase() == "irc legal" ||
              r.name.toLowerCase() == "irc medical wavier" ||
              r.name.toLowerCase() == "irc program administrator" ||
              r.name.toLowerCase() == "irc program manager" ||
              r.name.toLowerCase() == "irc qa" ||
              r.name.toLowerCase() == "irc rdtu" ||
              r.name.toLowerCase() == "irc service coordinator" ||
              r.name.toLowerCase() == "irc service coordinator org rights" ||
              r.name.toLowerCase() == "service writer"
          ).length > 0
      ) {
        setAbleToInitiateScanning(true);
      } else {
        setAbleToInitiateScanning(false);
      }
    }
  };
  let checkAccess = async (
    entityName: string,
    guid: string,
    userId: string
  ) => {
    var execute_RetrievePrincipalAccessInfo_Request = {
      // Parameters
      entity: { entityType: "systemuser", id: userId }, // entity
      ObjectId: { guid: guid }, // Edm.Guid
      EntityName: entityName, // Edm.String

      getMetadata: function () {
        return {
          boundParameter: "entity",
          parameterTypes: {
            entity: { typeName: "mscrm.systemuser", structuralProperty: 5 },
            ObjectId: { typeName: "Edm.Guid", structuralProperty: 1 },
            EntityName: { typeName: "Edm.String", structuralProperty: 1 },
          },
          operationType: 1,
          operationName: "RetrievePrincipalAccessInfo",
        };
      },
    };
    var response = await (props.context.webAPI as any)?.execute(
      execute_RetrievePrincipalAccessInfo_Request
    );
    if (response.ok) {
      var responseBody = await response.json();
      if (entityName == "contact") {
        var appendToAccess =
          JSON.parse(responseBody["AccessInfo"])
            .GrantedAccessRights.split(",")
            .filter(
              (it: any) =>
                it?.trim()?.toLowerCase() == "AppendToAccess".toLowerCase()
            ).length > 0;
        return appendToAccess;
      } else if (entityName == "neu_consumerdocument") {
        var accessToConsumerRecord =
          JSON.parse(responseBody["AccessInfo"])
            .GrantedAccessRights.split(",")
            .filter(
              (it: any) =>
                it?.trim()?.toLowerCase() == "CreateAccess".toLowerCase() &&
                it?.trim()?.toLowerCase() == "ReadAccess".toLowerCase() &&
                it?.trim()?.toLowerCase() == "WriteAccess".toLowerCase() &&
                it?.trim()?.toLowerCase() == "AppendAccess".toLowerCase()
            ).length > 0;
        return accessToConsumerRecord;
      }
    }
  };
  let onScanningDone=(isScanDone:boolean,err:string)=>{
    console.log('isScanDone',isScanDone); 
    setScanningDone(false);
    if(err){
      var alertStrings = { confirmButtonLabel: "ok", text: err, title: "Alert" };
      var alertOptions = { height: 120, width: 260 };
      props.context.navigation.openAlertDialog(alertStrings, alertOptions).then( function (success) {
        console.log("Alert dialog closed");
        location.reload()
    },
    function (error) {
        console.log(error.message);
    })
    }
   
  }
  let headerButtonClick=(types:HeaderButtonTypes)=>{
    console.log('types',types); 
    if(types==HeaderButtonTypes.InitiateScanning)
      acquireImage();
    else if(types==HeaderButtonTypes.loadDiskDocument)
      loadImage()
    else if(types==HeaderButtonTypes.ScanningCompleted)
      onDone();
  }
  return (
    <FluentProvider theme={webLightTheme}>    
   {isIRCDocumentRole&&
   <>   
   {!scanningDone&&
      <Header onButtonClick={headerButtonClick} disabled={!ableToInitiateScanning}
      showDisk={props.context.parameters.isDocumentLoadFromLocal.raw}    
      scanningViewExist={showScanningView}
      ></Header> 
    } 
     <div id={containerId} className={"display-" + showScanningView}></div>
      {scanningDone && (
        <ImageViewer DWObject={DWObject} context={props.context} onScanningDone={onScanningDone}></ImageViewer>
      )}
   </>
   } 
    
    </FluentProvider>
  );
};
