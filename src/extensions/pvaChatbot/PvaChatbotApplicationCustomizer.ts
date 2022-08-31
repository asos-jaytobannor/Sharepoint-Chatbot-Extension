import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import * as React from "react";  
import * as ReactDOM from "react-dom";  
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as strings from 'PvaChatbotApplicationCustomizerStrings';
import Chatbot from "./Chatbot";  

const LOG_SOURCE: string = 'PvaChatbotApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPvaChatbotApplicationCustomizerProperties {
  Bottom: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PvaChatbotApplicationCustomizer
  extends BaseApplicationCustomizer<IPvaChatbotApplicationCustomizerProperties> {
  private _bottomPlaceholder: PlaceholderContent | undefined;
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    // Wait for the placeholders to be created (or handle them being changed) and then
    // render.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    // Handling the bottom placeholder
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );

      // The extension should not assume that the expected placeholder is available.
      if (!this._bottomPlaceholder) {
        console.error("The expected placeholder (Bottom) was not found.");
        return;
      }
      const elem: React.ReactElement = React.createElement(Chatbot);
      ReactDOM.render(elem, this._bottomPlaceholder.domElement);
    }
  }
  private _onDispose(): void {
  }
}