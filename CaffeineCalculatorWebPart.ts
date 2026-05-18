import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import CaffeineCalculator from './components/CaffeineCalculator';
import { ICaffeineCalculatorProps } from './components/ICaffeineCalculatorProps';

export interface ICaffeineCalculatorWebPartProps {}

export default class CaffeineCalculatorWebPart extends BaseClientSideWebPart<ICaffeineCalculatorWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICaffeineCalculatorProps> = React.createElement(
      CaffeineCalculator,
      {}
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
