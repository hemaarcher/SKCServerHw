import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ArcherSerHwSysListWpWebPartStrings';
import ArcherSerHwSysListWp from './components/ArcherSerHwSysListWp';
import { IArcherSerHwSysListWpProps } from './components/IArcherSerHwSysListWpProps';

export interface IArcherSerHwSysListWpWebPartProps {
  description: string;
}

export default class ArcherSerHwSysListWpWebPart extends BaseClientSideWebPart<IArcherSerHwSysListWpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IArcherSerHwSysListWpProps> = React.createElement(
      ArcherSerHwSysListWp,
      {
        description: this.properties.description,
        spcontext:this.context,
      }
    );

    ReactDom.render(element, this.domElement);
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
