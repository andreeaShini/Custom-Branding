import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'CustomBrandingApplicationCustomizerStrings';

const LOG_SOURCE: string = 'CustomBrandingApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICustomBrandingApplicationCustomizerProperties {
  // This is an example; replace with your own property
  favicon: string;
  customcss: string;
  fa6css: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CustomBrandingApplicationCustomizer
  extends BaseApplicationCustomizer<ICustomBrandingApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const faviconUrl: string = this.properties.favicon;
    if (faviconUrl) {
      // inject the custom favicon
      let link = document.querySelector("link[rel*='icon']") as HTMLElement || document.createElement('link') as HTMLElement;
      link.setAttribute('type', 'image/x-icon');
      link.setAttribute('rel', 'shortcut icon');
      link.setAttribute('href', faviconUrl);
      document.getElementsByTagName('head')[0].appendChild(link);
    }

    const fa6css: string = this.properties.fa6css;
    if (fa6css) {
      // inject the custom style sheet
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      let customStyle: HTMLLinkElement = document.createElement("link");
      customStyle.rel = "stylesheet";
      customStyle.type = "text/css";  
      customStyle.href = fa6css;       
      head.insertAdjacentElement("beforeEnd", customStyle);
    }

    const cssUrl: string = this.properties.customcss;
    if (cssUrl) {
      // inject the custom style sheet
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      let customStyle: HTMLLinkElement = document.createElement("link");
      customStyle.rel = "stylesheet";
      customStyle.type = "text/css";
      customStyle.href = cssUrl;       
      head.insertAdjacentElement("beforeEnd", customStyle);
    }

    return Promise.resolve();
  }
}
