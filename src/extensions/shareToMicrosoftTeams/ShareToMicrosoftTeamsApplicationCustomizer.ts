import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName 
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'ShareToMicrosoftTeamsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'ShareToMicrosoftTeamsApplicationCustomizer';

require('./ShareToMicrosoftTeamsApplicationCustomizer.css');

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IShareToMicrosoftTeamsApplicationCustomizerProperties {
  // This is an example; replace with your own property
    showJustOnSitePages: string;
}


/** A Custom Action which can be run during execution of a Client Side Application */
export default class ShareToMicrosoftTeamsApplicationCustomizer
  extends BaseApplicationCustomizer<IShareToMicrosoftTeamsApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let showJustOnSitePages: string = this.properties.showJustOnSitePages; 

    this.context.application.navigatedEvent.add(this, this.initButtonValues);

    this.appendShareToTeamsScript();
    this._renderPlaceHolders(); 
    
    return Promise.resolve();
  }


  private _renderPlaceHolders(): void {

    // Handling the bottom placeholder
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this._onDispose });

      // The extension should not assume that the expected placeholder is available.
      if (!this._bottomPlaceholder) {
        console.error('The expected placeholder (Bottom) was not found.');
        return;
      }

      if (this.properties) {

        if (this._bottomPlaceholder.domElement) {
          this._bottomPlaceholder.domElement.innerHTML = `
          <div id="customShareTeamsBTN" class="teams-share-button" data-href="${document.location.href}"></div>`;
        }
      }
    }
    
  }

  private _onDispose(): void {
    console.log('Disposed Coments.');
  }

  private appendShareToTeamsScript(): void{
    //Add Share to Teams script to the page 
    var script   = document.createElement("script");
    script.type  = "text/javascript";
    script.src   = "https://teams.microsoft.com/share/launcher.js";
    document.body.appendChild(script);
  }

  private initButtonValues(): void{

    var requestURL = document.location.href;
    var waitEndNavigationInterval =setInterval(function(){
      if(requestURL!=document.location.href){
        var shareBTN = document.getElementById("customShareTeamsBTN");
        if(shareBTN){
          shareBTN.setAttribute("data-href", document.location.href); 
          eval('shareToMicrosoftTeams.renderButtons();');
        }    
      }
      clearInterval(waitEndNavigationInterval);
    },500);
    
  }


}


