import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-client-preview';

import * as strings from 'spfxSearchWithSliderPropertyStrings';
import SpfxSearchWithSliderProperty, { ISpfxSearchWithSliderPropertyProps } from './components/SpfxSearchWithSliderProperty';
import { ISpfxSearchWithSliderPropertyWebPartProps } from './ISpfxSearchWithSliderPropertyWebPartProps';

export default class SpfxSearchWithSliderPropertyWebPart extends BaseClientSideWebPart<ISpfxSearchWithSliderPropertyWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    const element: React.ReactElement<ISpfxSearchWithSliderPropertyProps> = React.createElement(SpfxSearchWithSliderProperty, {
      query: this.properties.query,
      count: this.properties.count,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      context: this.context
    });

    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
                PropertyPaneTextField('query', {
                  label: 'Query'
                }),
                PropertyPaneSlider('count', {
                  label: 'Count',
                  min: 1,
                  max: 50
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
