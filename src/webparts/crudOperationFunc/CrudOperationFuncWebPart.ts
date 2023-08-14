import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {sp} from '@pnp/sp/presets/all';

import * as strings from 'CrudOperationFuncWebPartStrings';
import CrudOperationFunc from './components/CrudOperationFunc';
import { ICrudOperationFuncProps } from './components/ICrudOperationFuncProps';

export interface ICrudOperationFuncWebPartProps {
  description: string;
}

export default class CrudOperationFuncWebPart extends BaseClientSideWebPart<ICrudOperationFuncWebPartProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
        spfxContext:this.context as any
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<ICrudOperationFuncProps> = React.createElement(
      CrudOperationFunc,
      {
        description: this.properties.description,
       context:this.context
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
