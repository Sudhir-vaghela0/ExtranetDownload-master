/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import styles from "./ExtranetDownloadButton.module.scss";
import type { IExtranetDownloadButtonProps } from "./IExtranetDownloadButtonProps";
import {
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  IComboBox,
  IComboBoxOption,
  PrimaryButton,
  Spinner,
} from "@fluentui/react";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import ComboBoxWithSearch from "../../../CommonComponents/ComboBoxWithSearch/ComboBoxWithSearch";
import {} from "@microsoft/sp-lodash-subset";
import { SharePointService } from "../../../Service/SharePointService";
import * as JSZip from "jszip";
import { saveAs } from "file-saver";
import * as _ from "lodash";

const dialogContentProps = {
  type: DialogType.largeHeader,
  title: "Download Langugae Pack",
  subText:
    "Required artifacts will be downloaded to your local system in zip format",
};

export interface IExtranetDownloadButtonState {
  hideDialog: boolean;
  selectedKeys: any[];
  options: IComboBoxOption[];
  selectedLangugaes: any[];
  isLoading: boolean;
}

export default class ExtranetDownloadButton extends React.Component<
  IExtranetDownloadButtonProps,
  IExtranetDownloadButtonState
> {
  constructor(props: IExtranetDownloadButtonProps) {
    super(props);
    this.state = {
      hideDialog: true,
      selectedKeys: [],
      options: [],
      selectedLangugaes: [],
      isLoading: false,
    };
  }

  async componentDidMount(): Promise<void> {
    const spService = new SharePointService();
    const allLangugaeCodes = await spService.getListItems(
      "LanguageCodes",
      "LanguageName"
    );
    if (allLangugaeCodes.results && allLangugaeCodes.results.length > 0) {
      this.setState({
        options: allLangugaeCodes.results.map((t: any) => {
          return { key: t.ID, text: t.LanguageName };
        }),
      });
    }
  }

  public toggleHideDialog = (): void => {
    this.setState({ hideDialog: true });
  };

  public captureCurrentPageAsPDF = (): Promise<Blob> => {
    return new Promise<Blob>((resolve, reject) => {
      // document.querySelector('div[data-automation-id="contentScrollRegion"] > div') as any
      html2canvas(
        document.getElementsByClassName("CanvasComponent")[0] as any,
        {
          scale: 2,
          useCORS: true,
          allowTaint: true,
        }
      )
        .then((canvas) => {
          const imgData = canvas.toDataURL("image/png");
          const imgWidth = 210;
          const pageHeight = 295;
          const imgHeight = (canvas.height * imgWidth) / canvas.width;
          let heightLeft = imgHeight;
          let position = 10;
          const doc = new jsPDF("p", "mm");
          doc.addImage(imgData, "PNG", 0, 0, imgWidth, imgHeight);
          heightLeft -= pageHeight;
          while (heightLeft >= 0) {
            position += heightLeft - imgHeight; // top padding for other pages
            doc.addPage();
            doc.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
            heightLeft -= pageHeight;
          }
          const pdfBlob = doc.output("blob");
          resolve(pdfBlob);
        })
        .catch(reject);
    });
  };

  public onChange = (
    event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption,
    index?: number,
    value?: string
  ): void => {
    const selected = option?.selected;
    if (option) {
      this.setState({
        selectedKeys: selected
          ? [...this.state.selectedKeys, option!.key as string]
          : this.state.selectedKeys.filter((k) => k !== option!.key),
        selectedLangugaes: selected
          ? [...this.state.selectedLangugaes, option!.text as string]
          : this.state.selectedLangugaes.filter((k) => k !== option!.text),
      });
    }
  };

  private fetchDocumentsFromLibrary = async (
    libraryName: string,
    langCode: number
  ): Promise<Array<{ fileName: string; fileContent: Blob }>> => {
    const spService = new SharePointService();
    const files = await spService.getFilteredListItems(
      libraryName,
      `LangCode/Id eq ${langCode}`,
      "FileLeafRef",
      ["LangCode/Id", "FileLeafRef", "FileRef"],
      ["LangCode"]
    );

    const documents: Array<{ fileName: string; fileContent: Blob }> = [];
    for (const file of files.results) {
      const fileContent = await spService.getFileBlob(file.FileRef);
      documents.push({ fileName: file.FileLeafRef, fileContent });
    }

    return documents;
  };

  private saveUserInfo = async (): Promise<void> => {
    const spService = new SharePointService();
    await spService.saveItemToList("DownloadInformation", {
      SelectedLanguages: this.state.selectedLangugaes.join(";"),
      Title: document.title,
    });
  };

  private downloadPageAndDocumentsAsZip = async (): Promise<void> => {
    this.setState({ isLoading: true });
    await this.saveUserInfo();
    const zip = new JSZip();
    const currentPagePDF = await this.captureCurrentPageAsPDF();
    zip.file(`${document.title}.pdf`, currentPagePDF);
    console.log(this.state.selectedKeys);
    let documents: Array<{ fileName: string; fileContent: Blob }> = [];
    await Promise.all(
      this.state.selectedKeys.map(async (t) => {
        const temp = await this.fetchDocumentsFromLibrary("Documents", t);
        documents = _.merge(documents, temp);
      })
    );
    for (const doc of documents) {
      zip.file(doc.fileName, doc.fileContent);
    }

    zip
      .generateAsync({ type: "blob" })
      .then((content) => {
        this.setState({ hideDialog: true, isLoading: false });
        saveAs(content, `${this.props.context.pageContext.web.title}_${this.state.selectedLangugaes.length > 2 ? 'multi' : this.state.selectedLangugaes[0]}`);
      })
      .catch((error) => console.error("Error generating ZIP:", error));
  };

  public render(): React.ReactElement<IExtranetDownloadButtonProps> {
    return (
      <div className={styles.extranetDownloadButton}>
        <PrimaryButton
          text="Download Language Pack offline"
          onClick={() => this.setState({ hideDialog: false })}
          className={styles.btn}
        />
        <Dialog
          hidden={this.state.hideDialog}
          onDismiss={this.toggleHideDialog}
          dialogContentProps={dialogContentProps}
          modalProps={{ isBlocking: true }}
        >
          {this.state.isLoading && (
            <Spinner
              label="Please Wait, while we fetch the data ..."
              ariaLive="assertive"
              labelPosition="right"
            />
          )}
          {!this.state.isLoading && (
            <>
              <ComboBoxWithSearch
                filteredOptions={this.state.options}
                onChange={this.onChange}
                label="Select a language"
              />
              <DialogFooter>
                <PrimaryButton
                  onClick={this.downloadPageAndDocumentsAsZip}
                  text="Download"
                />
                <DefaultButton onClick={this.toggleHideDialog} text="Cancel" />
              </DialogFooter>
            </>
          )}
        </Dialog>
      </div>
    );
  }
}
