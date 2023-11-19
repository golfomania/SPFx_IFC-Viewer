import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import styles from "./IfCviewerWebPart.module.scss";
import { SPHttpClient } from "@microsoft/sp-http";

export interface IIfCviewerWebPartProps {
  description: string;
}

export default class IfCviewerWebPart extends BaseClientSideWebPart<IIfCviewerWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `
    <div id="myIFCviewer" class=${styles.spfxViewer}></div>
    `;
  }

  protected onInit() {
    setTimeout(async () => {
      await import(
        // @ts-ignore
        /* webpackIgnore: true */ "https://golfomania.github.io/SPFx_IFC-Viewer/dist/assets/index-d85e628e.js"
      );

      window.dispatchEvent(new Event("resize"));

      await this.loadFirstFile();
    }, 1000);

    return new Promise<void>((resolve) => {
      console.log(styles);
      resolve();
    });
  }

  protected async loadFirstFile() {
    const docFiles = "_api/web/lists/GetByTitle('ifcFiles')/Files";
    const baseUrl = this.context.pageContext.web.absoluteUrl;
    const url = `${baseUrl}/${docFiles}`;

    const http = this.context.spHttpClient;
    const config = SPHttpClient.configurations.v1;
    const response = await http.get(url, config);
    const documents = await response.json();

    console.log(documents);

    if (documents.value.length) {
      const firstDocument = documents.value[0].Url;
      const fetched = await fetch(firstDocument);
      const buffer = await fetched.arrayBuffer();
      const data = new Uint8Array(buffer);

      const event = new CustomEvent("loadIFC", {
        detail: {
          name: "loadIFC",
          payload: {
            name: "example",
            buffer: data,
          },
        },
      });
      window.dispatchEvent(event);
    }
  }
}
