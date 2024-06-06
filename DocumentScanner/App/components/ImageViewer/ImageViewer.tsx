import * as React from "react";
import {
  ArrowNext16Regular,
  ArrowPrevious16Regular,
  Search12Regular,
} from "@fluentui/react-icons";
import {
  Field,
  Input,
  Switch,
  Textarea,
  Card,
  CardFooter,
  CardPreview,
  Button,
  ButtonProps,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from "@fluentui/react-components";
import { WebTwain } from "dwt/dist/types/WebTwain";
import { useEffect, useState } from "react";
import Dynamsoft from "dwt";
import { PDFDocument } from "pdf-lib";

import { consumerForm } from "../../../models/IconsumerForm";

import { IInputs } from "../../../generated/ManifestTypes";
const defaultFormValues = {
  carryDtToSqntPage: {
    value: true,
  },
  startNewDocument: {
    value: true,
  },
  consumer: {
    value: "",
  },
  documentType: {
    value: "",
  },
  documentName: {
    value: "",
  },
  documentDesc: {
    value: "",
  },
  image: "",
};
const SearchButton: React.FC<ButtonProps> = (props) => {
  return (
    <Button
      {...props}
      appearance="transparent"
      icon={<Search12Regular />}
      size="small"
    />
  );
};
export const ImageViewer = (props: {
  DWObject: WebTwain;
  context: ComponentFramework.Context<IInputs>;
  onScanningDone: (isScanDone: boolean, err: string) => void;
}) => {
  const [currentPage, setCurrentPage] = useState<number>(0);
  const [currentImage, setCurrentImage] = useState<string>("");
  const [consumerFormValues, setConsumerFormValues] = useState<consumerForm>({
    ...defaultFormValues,
    pageNo: currentPage,
  });
  const [counsumerDocuments, setcounsumerDocuments] = useState<consumerForm[]>([
    consumerFormValues,
  ]);
  const [scanningDone, SetscanningDone] = useState(false);
  useEffect(() => {
    setConsumerFormValues((values) => ({
      ...values,
      startNewDocument: { value: currentPage == 0 },
      consumer: {
        value: props.context.parameters.primaryField.raw,
        id: (props.context as any).page.entityId,
      },
    }));
    props.DWObject.ConvertToBase64(
      [currentPage],
      Dynamsoft.DWT.EnumDWT_ImageType.IT_PNG,
      function (result, indices, type) {
        let data = result.getData(0, result.getLength());
        setCurrentImage("data:image/png; base64," + data);
        setConsumerFormValues((pre) => ({ ...pre, image: data }));
      },
      function (errorCode, errorString) {
        console.log(errorString);
      }
    );
  }, [currentPage]);

  const onClickOnNext = () => {
    let isFormValid = formValidate(consumerFormValues);
    if (!isFormValid) {
      let nextPage = currentPage + 1;
      let lastObj =
        counsumerDocuments.length > 0
          ? counsumerDocuments[counsumerDocuments.length - 1]
          : null;
      let isPageValueAlreadyExist=counsumerDocuments.find(it=>it.pageNo==nextPage);
      if(!isPageValueAlreadyExist){
        setConsumerFormValues((x) => ({
          ...defaultFormValues,
          pageNo: nextPage,
          consumer: x.consumer,
          startNewDocument: x.startNewDocument,
          carryDtToSqntPage: x.carryDtToSqntPage,
          documentName: x.documentName,
          documentDesc: x.documentDesc,
          documentType: lastObj?.carryDtToSqntPage.value
            ? x.documentType
            : {
                value: "",
              },
        }));
      }else{
        setConsumerFormValues((x) => ({
          ...x,
          ...isPageValueAlreadyExist
          // pageNo: nextPage           
        }));
      }
     
      console.log(
        consumerFormValues,
        " var documents=Array.from(counsumerDocuments);"
      );
      upsertDocument(consumerFormValues);
      setCurrentPage(nextPage);
    }
  };
  const onClickOnGoToLast = async () => {
    let isFormValid = formValidate(consumerFormValues);
    if (!isFormValid) {
      //let totalItems=props.DWObject.HowManyImagesInBuffer;
      let lastObj1 =
        counsumerDocuments.length > 0
          ? counsumerDocuments[counsumerDocuments.length - 1]
          : null;
      setConsumerFormValues((x) => ({
        ...defaultFormValues,
        pageNo: currentPage,
        consumer: x.consumer,
        startNewDocument: x.startNewDocument,
        carryDtToSqntPage: x.carryDtToSqntPage,
        documentName: x.documentName,
        documentDesc: x.documentDesc,
        documentType: lastObj1?.carryDtToSqntPage.value
          ? x.documentType
          : {
              value: "",
            },
      }));
      upsertDocument(consumerFormValues, true);
    }
  };
  const getImage = (currentPage: number) => {
    return new Promise(function (resolve, reject) {
      // getUserData(userId, resolve, reject);
      props.DWObject.ConvertToBase64(
        [currentPage],
        Dynamsoft.DWT.EnumDWT_ImageType.IT_PNG,
        function (result, indices, type) {
          let data = result.getData(0, result.getLength());
          resolve(data);
        },
        function (errorCode, errorString) {
          console.log(errorString);
          reject(errorString);
        }
      );
    });
  };
  const upsertDocument = async (
    consumerFormValue: consumerForm,
    forGoToLastPage = false
  ) => {
    var documents = Array.from(counsumerDocuments);
    let isIndex = documents.findIndex(
      (d) => d.pageNo == consumerFormValue.pageNo
    );
    if (isIndex > -1) {
      documents[isIndex] = consumerFormValue;
    } else documents.push(consumerFormValue);
    setcounsumerDocuments(documents);
    if (forGoToLastPage) {
      let totalItems = props.DWObject.HowManyImagesInBuffer;
      let nextPage = currentPage + 1;
      var items = [...documents];
      while (nextPage < totalItems - 1) {
        let lastObj =
          documents.length > 0 ? documents[documents.length - 1] : null;
        let nextItem = {
          ...defaultFormValues,
          pageNo: nextPage,
          consumer: consumerFormValues.consumer,
          startNewDocument: { value: false },
          carryDtToSqntPage: consumerFormValues.carryDtToSqntPage,
          documentName: consumerFormValues.documentName,
          documentDesc: consumerFormValues.documentDesc,
          documentType: lastObj?.carryDtToSqntPage.value
            ? consumerFormValues.documentType
            : {
                value: "",
              },
        };
        var imageData: string = (await getImage(nextPage)) as any;
        nextItem.image = imageData;
        items.push(nextItem);
        nextPage++;
      }
      console.log("value", items);
      setcounsumerDocuments(items);
      if(items.length>(nextPage-1))
        {
         var currentPageInfo= items[nextPage-1];
         if(currentPageInfo!=null){
          var imageData1: string = (await getImage(nextPage)) as any;
           
          setConsumerFormValues((x) => ({
            ...currentPageInfo,
            pageNo:nextPage ,
            image:imageData1              
          }));
         }       
        }
      setCurrentPage(nextPage);
    }
  };
  useEffect(() => {
    if (
      counsumerDocuments.length ==
      props.DWObject.GetDocumentInfoList()?.[0].imageIds.length
    ) { 
      mergeToPDF();
    }
  }, [counsumerDocuments]);
  const mergeToPDF = async () => {
    type documentGroup = { documentCount: any; items: consumerForm[] };
    var groupDocuments: documentGroup[] = [];
    var documentCount = 1;
    for (let i = 0; i < counsumerDocuments.length; i++) {
      const element = counsumerDocuments[i];
      if (element.startNewDocument.value) {
        groupDocuments.push({
          documentCount: documentCount,
          items: [element],
        });
        documentCount++;
      } else {
        var goupDocumentFound = groupDocuments.find(
          (it) => it.documentCount === documentCount - 1
        );
        if (goupDocumentFound) {
          goupDocumentFound.items.push(element);
        }
      }
    }
    console.log("groupDocuments", groupDocuments);
    type consumerDocument = {
      consumer: any;
      documentType: any;
      name: string;
      description: string;
      source: string;
      noOfPages: number;
      UploadDate: Date;
      document: string;
    };

    try {
      var xrm = (window as any).Xrm;
      if (xrm) xrm.Utility.showProgressIndicator("Processing....");
      for (let i = 0; i < groupDocuments.length; i++) {
        const group = groupDocuments[i];
        if (group.items.length > 0) {
          var mergedPDf = await mergeDocuments(group.items);
          let fItem = group.items[0];
          let documentName = "";
          var documentNameValue = fItem.documentName.value?.trim();
          if (documentNameValue) documentName = documentNameValue;
          var data = {
            "neu_Consumer@odata.bind": "/contacts(" + fItem.consumer.id + ")",
            "neu_DocumentType@odata.bind":
              "/neu_documenttypes(" +
              fItem.documentType.id?.replace("}", "").replace("{", "") +
              ")",
            neu_name: documentName,
            neu_description: fItem.documentDesc.value?.trim(),
            neu_documentsource: 288500002,
            neu_numberofpages: group.items.length,
            neu_uploaddate: new Date(),
          };
          console.log("data", data);
          //tempary comment
          var createdConsumerId = await props.context.webAPI.createRecord(
            "neu_consumerdocument",
            data
          );
          await createNote(
            createdConsumerId.id,
            createdConsumerId.entityType,
            documentName + ".pdf",
            mergedPDf
          );
        }
      }
      console.log("PDF creation successful!");
      SetscanningDone(true);
      props.onScanningDone(true, "Document(s) scanned successfully.");
    } catch (error) {
      console.error("Error merging PNGs to PDF:", error);
      console.log("Error merging PNGs to PDF");
    }
    if (xrm) xrm.Utility.closeProgressIndicator();
  };

  let createNote = async (
    entityId: string,
    entityType: string,
    fileName: string,
    fileContentBase64: string
  ) => {
    var note = {
      "objectid_neu_consumerdocument@odata.bind":
        "/neu_consumerdocuments(" + entityId + ")",
      subject: fileName,
      documentbody: fileContentBase64,
      filename: fileName,
    };

    var result = await props.context.webAPI.createRecord("annotation", note);
    return result.id;
  };
  const mergeDocuments = async (counsumerDocuments: consumerForm[]) => {
    const pdfDoc = await PDFDocument.create();
    for (let i = 0; i < counsumerDocuments.length; i++) {
      const pngBase64 = counsumerDocuments[i].image;
      if (pngBase64) {
        const pngArrayBuffer = Uint8Array.from(atob(pngBase64), (c) =>
          c.charCodeAt(0)
        );

        const pngImage = await pdfDoc.embedPng(pngArrayBuffer);
        const pngDims = pngImage.scale(0.5);
        const page = pdfDoc.addPage();

        const aspectRatio = pngDims.width / pngDims.height;
        const imageWidth = Math.min(
          page.getWidth(),
          page.getHeight() * aspectRatio
        );
        const imageHeight = imageWidth / aspectRatio;
        const x = (page.getWidth() - imageWidth) / 2;
        const y = (page.getHeight() - imageHeight) / 2;

        page.drawImage(pngImage, {
          x: x,
          y: y,
          width: imageWidth,
          height: imageHeight,
        });
      }
    }
    const pdfBytes = await pdfDoc.save();
    const pdfBase64 = arrayBufferToBase64(pdfBytes);
    return pdfBase64;
  };
  const arrayBufferToBase64 = (buffer: any) => {
    let binary = "";
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
  };
  const onClickOnPrevious = () => {
    let previousPage = currentPage - 1;
    setCurrentPage(previousPage);
    let previousDocument = counsumerDocuments.find(
      (x) => x.pageNo == previousPage
    );
    if (previousDocument) setConsumerFormValues({ ...previousDocument });
  };
  const formChange = (name: string, value?: any) => {
    var values = { ...consumerFormValues };
    if (name == "startNewDocument" && !value && counsumerDocuments.length > 0) {
      var lastObj = counsumerDocuments[counsumerDocuments.length - 1];
      setConsumerFormValues((values) => ({
        ...values,
        [name]: {
          value: value,
        },
        documentType: lastObj?.documentType,
        documentName: lastObj?.documentName,
        documentDesc: lastObj?.documentDesc,
      }));
      return;
    } else if (
      name == "carryDtToSqntPage" &&
      values.startNewDocument.value &&
      counsumerDocuments.length > 0
    ) {
      var lastObjDocument = counsumerDocuments[counsumerDocuments.length - 1];
      if (value) {
        setConsumerFormValues((values) => ({
          ...values,
          [name]: {
            value: value,
          },
          documentType: lastObjDocument.documentType,
        }));
      } else {
        setConsumerFormValues((values) => ({
          ...values,
          [name]: {
            value: value,
          },
          documentType: {
            value: "",
            id: "",
          },
        }));
      }

      return;
    }

    setConsumerFormValues((values) => ({
      ...values,
      [name]: {
        value: value,
      },
    }));
  };
  const onSubmit = () => {
    let isFormValid = formValidate(consumerFormValues);
    if (!isFormValid) {
      upsertDocument(consumerFormValues);
    }
  };
  const formValidate = (consumerFormValues: consumerForm) => {
    let isNullOrWhitespace = (input: string) => {
      return !input || !input.trim();
    };
    var newFormAfterValid = {
      ...consumerFormValues,
      consumer: {
        ...consumerFormValues.consumer,
        error: isNullOrWhitespace(consumerFormValues.consumer.value)
          ? "Consumer is required."
          : "",
      },
      documentType: {
        ...consumerFormValues.documentType,
        error: isNullOrWhitespace(consumerFormValues.documentType.value)
          ? "Document type is required."
          : "",
      },
      documentName: {
        ...consumerFormValues.documentName,
        error:
          consumerFormValues.documentName.value.length > 100
            ? "100 characters allowed."
            : "",
      },
      documentDesc: {
        ...consumerFormValues.documentDesc,
        error:
          consumerFormValues.documentDesc.value.length > 8000
            ? "8000 characters allowed."
            : "",
      },
    };
    setConsumerFormValues(newFormAfterValid);
    let isError =
      Object.keys(newFormAfterValid).findIndex(
        (key) => (newFormAfterValid as any)[key].error
      ) > 0;
    console.log("isError", isError);
    return isError;
  };
  const searchClick = (e: any) => {
    console.log("search clicked", consumerFormValues.documentType.value);
    var lookupOptions = {
      defaultEntityType: "neu_documenttype",
      entityTypes: ["neu_documenttype"],
      searchText: consumerFormValues.documentType.value,
      disableMru: true,
    };
    props.context.utils.lookupObjects(lookupOptions).then((it) => {
      var lookRef = it?.[0];
      if (lookRef)
        setConsumerFormValues((prevalue) => ({
          ...prevalue,
          documentType: { value: lookRef.name, id: lookRef.id },
        }));
      else
        setConsumerFormValues((prevalue) => ({
          ...prevalue,
          documentType: { value: "", id: "" },
        }));
    });
  };
  return (
    <div>
      <Card>
        <div
          style={{
            display: !scanningDone ? "flex" : "none",
            justifyContent: "space-between",
          }}
        >
          <div>
            <CardPreview>
              <img src={currentImage} alt="Presentation Preview" />
            </CardPreview>
          </div>
          <div className="Fields">
            <Field label="Carry over Document Type to subsequent pages/documents">
              <Switch
                checked={consumerFormValues.carryDtToSqntPage.value}
                onChange={(ev: React.ChangeEvent<HTMLInputElement>) =>
                  formChange("carryDtToSqntPage", ev.target.checked)
                }
                label={
                  consumerFormValues.carryDtToSqntPage.value ? "Yes" : "No"
                }
              />
            </Field>
            <Field label="Start New Document">
              <Switch
                checked={consumerFormValues.startNewDocument.value}
                onChange={(ev: React.ChangeEvent<HTMLInputElement>) =>
                  formChange("startNewDocument", ev.target.checked)
                }
                label={consumerFormValues.startNewDocument.value ? "Yes" : "No"}
                disabled={currentPage == 0}
              />
            </Field>
            <Field
              label="Consumer"
              required
              validationMessage={consumerFormValues.consumer.error}
            >
              <Input
                placeholder="Enter Consumer"
                value={consumerFormValues.consumer.value}
                onChange={(ev: React.ChangeEvent<HTMLInputElement>) =>
                  formChange("consumer", ev.target.value)
                }
                disabled={true}
              />
            </Field>
            <Field
              label="Document Type"
              required
              validationMessage={consumerFormValues.documentType.error}
            >
              <Input
                placeholder="Enter Document Type"
                value={consumerFormValues.documentType.value}
                onChange={(ev: React.ChangeEvent<HTMLInputElement>) =>
                  formChange("documentType", ev.target.value)
                }
                disabled={!consumerFormValues.startNewDocument.value}
                contentAfter={
                  <SearchButton
                    aria-label="Search..."
                    onClick={searchClick}
                    disabled={!consumerFormValues.startNewDocument.value}
                  />
                }
              />
            </Field>
            <Field
              label="Document Name"
              validationMessage={consumerFormValues.documentName.error}
            >
              <Input
                placeholder="Enter Document Name"
                value={consumerFormValues.documentName.value}
                onChange={(ev: React.ChangeEvent<HTMLInputElement>) =>
                  formChange("documentName", ev.target.value)
                }
                disabled={!consumerFormValues.startNewDocument.value}
                maxLength={100}
              />
            </Field>
            <Field
              label="Document Description"
              validationMessage={consumerFormValues.documentDesc.error}
            >
              <Textarea
                placeholder="Enter Document Description"
                value={consumerFormValues.documentDesc.value}
                onChange={(ev: React.ChangeEvent<HTMLTextAreaElement>) =>
                  formChange("documentDesc", ev.target.value)
                }
                disabled={!consumerFormValues.startNewDocument.value}
                maxLength={8000}
              />
            </Field>
            <br />
            <CardFooter>
              <div>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                  }}
                >
                  <Button
                    icon={<ArrowPrevious16Regular />}
                    appearance="primary"
                    disabled={currentPage == 0}
                    onClick={onClickOnPrevious}
                  >
                    Previous
                  </Button>
                  {!(
                    currentPage ==
                    props.DWObject.GetDocumentInfoList()?.[0].imageIds.length -
                      1
                  ) && (
                    <>
                      <Button
                        style={{ marginLeft: "10px" }}
                        onClick={onClickOnNext}
                        appearance="primary"
                        icon={<ArrowNext16Regular />}
                      >
                        Next
                      </Button>
                    </>
                  )}
                </div>
                {!(
                  currentPage ==
                  props.DWObject.GetDocumentInfoList()?.[0].imageIds.length - 1
                ) && (
                  <div style={{ margin: "10px", textAlign: "center" }}>
                    <Button
                      onClick={onClickOnGoToLast}
                      appearance="primary"
                      icon={<ArrowNext16Regular />}
                    >
                      Go to Last Page
                    </Button>
                  </div>
                )}
              </div>

              {currentPage ==
                props.DWObject.GetDocumentInfoList()?.[0].imageIds.length -
                  1 && (
                <Button appearance="primary" onClick={() => onSubmit()}>
                  Submit
                </Button>
              )}
            </CardFooter>
          </div>
        </div>
      </Card>
    </div>
  );
};
