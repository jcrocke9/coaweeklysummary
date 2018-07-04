import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'CoaWeeklySummaryWebPartStrings';
import CoaWeeklySummary from './components/CoaWeeklySummary';
import { ICoaWeeklySummaryProps } from './components/ICoaWeeklySummaryProps';

export interface ICoaWeeklySummaryWebPartProps {
  description: string;
}

export default class CoaWeeklySummaryWebPart extends BaseClientSideWebPart<ICoaWeeklySummaryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICoaWeeklySummaryProps > = React.createElement(
      CoaWeeklySummary,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
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
                }),
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel,
                  value: "Weekly Summary"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
