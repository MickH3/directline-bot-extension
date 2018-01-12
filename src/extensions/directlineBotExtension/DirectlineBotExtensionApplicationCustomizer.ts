import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';


import * as strings from 'DirectlineBotExtensionApplicationCustomizerStrings';
import { escape } from '@microsoft/sp-lodash-subset';

//npm install botframework-webchat
import { App, DirectLine } from 'botframework-webchat';
import { ChatProps } from 'botframework-webchat/built/Chat';
require("../../../node_modules/botframework-webchat/botchat.css");
require("./InlineCSS.css");

//npm install --save jquery@2
//npm install --save @types/jquery@2
import * as jQuery from 'jquery';

//npm install @microsoft/sp-office-ui-fabric-core --save-dev
//@import '~office-ui-fabric-core/dist/sass/Fabric.scss'; 
import styles from './OfficeUI.module.scss';


const LOG_SOURCE: string = 'DirectlineBotExtensionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDirectlineBotExtensionApplicationCustomizerProperties {
  DirectLineSecret: string,
  BotId: string,
  BotName: string
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class DirectlineBotExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<IDirectlineBotExtensionApplicationCustomizerProperties> {

  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _botConnection: DirectLine | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    this._renderPlaceHolders();

    return Promise.resolve();
  }

  public _openChat(): void{
    //When called as a click event, 'this' is the HTML element, not my class.  Grab the reference to my class stored in the <div id="bot"> element.
    let _that : DirectlineBotExtensionApplicationCustomizer = jQuery("#bot").data("ref");

    jQuery("#chatOpener").fadeOut("fast", function () {
      jQuery("#chatWindow").fadeIn("fast", function(){
        if(!_that._botConnection){
          _that._botConnection = new DirectLine({
            secret: _that.properties.DirectLineSecret
          });

          App({
            user: { id: _that.context.pageContext.user.loginName + "." + _that.properties.BotName, name: _that.context.pageContext.user.displayName },
            botConnection: _that._botConnection,
            bot: { id: _that.properties.BotId, name: _that.properties.BotName},
            sendTyping: true
          }, document.getElementById("bot"));
        }
      });
    });
  }

  public _closeChat(): void{
    jQuery("#chatWindow").fadeOut("fast", function(){
      jQuery("#chatOpener").fadeIn("fast");
    });
  }

  private _renderPlaceHolders(): void{
    console.log('DirectlineBotExtensionApplicationCustomizer._renderPlaceHolders()');
    console.log('Available placeholders: ', this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(', '));
    

    if(!this._bottomPlaceholder){
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom, { onDispose: this._onDispose }
      );

      if(!this._bottomPlaceholder){
        console.error('The expected placeholder (Bottom) was not found.');
        return;
      }
    

      if(this.properties){
        if(this._bottomPlaceholder.domElement){
          this._bottomPlaceholder.domElement.innerHTML = `
          <div>
            <div id="chatOpener" class="ms-bgColor-themeDark chat-opener">
                <a class="b" href="javascript:">
                    <i class="${styles['ms-Icon']} ${styles['ms-Icon--CommentPrevious']}"></i>
                    <span>${this.properties.BotName}</span>
                </a>
            </div>
            <div id="chatWindow" class="chat-window" style="display:none;">
                <div class="ms-bgColor-themeDark heading">
                    <a href="javascript:" id="chatCloser" title="Close chat window"><i class="${styles['ms-Icon']} ${styles['ms-Icon--Cancel']}"></i></a> ${this.properties.BotName}
                </div>
                <div id="bot"></div>
            </div>
          </div>
          `;

          document.getElementById("chatOpener").onclick = this._openChat;//OnClick event assignment
          document.getElementById("chatCloser").onclick = this._closeChat;//OnClick event assignment

          //Possibly due to how we are assigning click events, 'this' param in our onclick functions isn't the class.
          //Assign a ref to the class we can retrieve later
          jQuery("#bot").data("ref", this);
        }
      }
    }
  }

  private _onDispose(): void {
    console.log('[DirectlineBotExtensionApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
