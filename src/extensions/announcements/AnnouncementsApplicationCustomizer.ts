import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'AnnouncementsApplicationCustomizerStrings';
import styles from "./AnnouncementsApplicationCustomizer.module.scss";
const LOG_SOURCE: string = 'AnnouncementsApplicationCustomizer';
import { sp, Web } from "@pnp/sp";
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAnnouncementsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AnnouncementsApplicationCustomizer
  extends BaseApplicationCustomizer<IAnnouncementsApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;
  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    sp.setup({
      spfxContext: this.context
    });
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

  }

  private async _renderPlaceHolders() {
    if(!this._topPlaceholder){
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
    }

    if(this._topPlaceholder){
      this._topPlaceholder.domElement.innerHTML = `<div class="${styles.announcements}">Loading...</div>`;

      var announcement = await this.getAnnouncement();
      if(announcement){
        this._topPlaceholder.domElement.innerHTML = `<div class="${styles.announcements}">${announcement.Title}</div>`;
      }else{
        this._topPlaceholder.domElement.innerHTML = ``;
      }
      
    }
  }

  private async getAnnouncement(){
    var web = new Web("/sites/apps/");
    var announcments = await web.lists.getByTitle("Announcements").items.orderBy("Modified", false).top(1).get();

    return announcments[0];
  }
}
