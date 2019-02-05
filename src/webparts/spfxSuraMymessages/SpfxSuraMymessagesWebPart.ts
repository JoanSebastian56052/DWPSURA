import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxSuraMymessagesWebPartStrings';
import SpfxSuraMymessages from './components/SpfxSuraMymessages';
import { ISpfxSuraMymessagesProps } from './components/ISpfxSuraMymessagesProps';

import { MSGraphClient } from '@microsoft/sp-http';

export interface ISpfxSuraMymessagesWebPartProps {
  description: string;
}

export default class SpfxSuraMymessagesWebPart extends BaseClientSideWebPart<ISpfxSuraMymessagesWebPartProps> {

  public render(): void {
    this.context.msGraphClientFactory.getClient()
      .then((client: MSGraphClient): void => {
        const element: React.ReactElement<ISpfxSuraMymessagesProps > = React.createElement(
          SpfxSuraMymessages,
          {
            graphClient: client
          }
        );
        ReactDom.render(element, this.domElement);
      })
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
