import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import { Dialog } from "@microsoft/sp-dialog";

import * as strings from "GoogleAnalyticsForSpoApplicationCustomizerStrings";

const LOG_SOURCE: string = "GoogleAnalyticsForSpoApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGoogleAnalyticsForSpoApplicationCustomizerProperties {
  // This is an example; replace with your own property
  trackingID: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GoogleAnalyticsForSpoApplicationCustomizer extends BaseApplicationCustomizer<IGoogleAnalyticsForSpoApplicationCustomizerProperties> {
  /** Current SharePoint page navigation
   * @private
   */
  private currentPage = "";

  /** Statement for trigger ony once the initialization of analytics script
   * @private
   */
  private isInitialLoad = true;

  /** Get Current page URL
   * @returns URL of the current Page
   * @private
   */
  private getFreshCurrentPage(): string {
    return window.location.pathname + window.location.search;
  }

  /** Update current page navigation
   * @private
   */
  private updateCurrentPage(): void {
    this.currentPage = this.getFreshCurrentPage();
  }

  /** Navigation and search event
   * @private
   */
  private navigatedEvent(): void {
    let trackingID: string = this.properties.trackingID;
    if (!trackingID) {
      Log.info(LOG_SOURCE, `${strings.MissingID}`);
    } else {
      Log.info(LOG_SOURCE, `Tracking Site ID: ${trackingID}`);
      const navigatedPage = this.getFreshCurrentPage();

      if (this.isInitialLoad) {
        Log.info(LOG_SOURCE, `Initial load`);
        this.realInitialNavigatedEvent(trackingID);
        this.updateCurrentPage();
        this.isInitialLoad = false;
      } else if (!this.isInitialLoad && navigatedPage !== this.currentPage) {
        Log.info(LOG_SOURCE, `Not initial load`);
        this.realNavigatedEvent(trackingID);
        this.updateCurrentPage();
      }
    }
  }

  /** Inital Page load - init analytics
   * @param trackingID Google Analytics Tracking Site ID
   * @private
   */
  private realInitialNavigatedEvent(trackingID: string): void {
    Log.info(LOG_SOURCE, `Tracking full page load...`);

    var gtagScript = document.createElement("script");
    gtagScript.type = "text/javascript";
    gtagScript.src = `https://www.googletagmanager.com/gtag/js?id=${trackingID}`;
    gtagScript.async = true;
    document.head.appendChild(gtagScript);

    eval(`
           window.dataLayer = window.dataLayer || [];
           function gtag(){dataLayer.push(arguments);}
           gtag('js', new Date());
           gtag('config',  '${trackingID}');
         `);
  }

  /** Partial Page load
   * @param trackingID Google Analytics Tracking Site ID
   * @private
   */
  private realNavigatedEvent(trackingID: string): void {
    Log.info(LOG_SOURCE, `Tracking partial page load...`);

    eval(`
         if(ga) {
           ga('create', '${trackingID}', 'auto');
           ga('set', 'page', '${this.getFreshCurrentPage()}');
           ga('send', 'pageview');
         }
         `);
  }

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized Google Analytics`);
    /* This event is triggered when user performed a search from the header of SharePoint */
    this.context.placeholderProvider.changedEvent.add(
      this,
      this.navigatedEvent
    );
    /* This event is triggered when user navigate between the pages */
    this.context.application.navigatedEvent.add(this, this.navigatedEvent);

    return Promise.resolve();
  }
}
