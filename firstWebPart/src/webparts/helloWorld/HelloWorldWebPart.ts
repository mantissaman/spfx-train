import { Version, DisplayMode, Environment, EnvironmentType, Log } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

export interface IHelloWorldWebPartProps {
  description: string;
  myContinent: string;
  numContinentsVisited: number;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {

    if (!this.renderedOnce) {
      // SPComponentLoader.loadScript('https://.../jquery.min.js', jQuery)
      //   .then($: any): void ={
      //   this.jQuery = $;
      //   SPComponentLoader.loadCss('https://.../jqueryui.css');
      // }
    }

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, "I am loading please wait");
    setTimeout(() => {
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    }, 1000);
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorld}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">This is our first webpart.</p>
              <p class="${ styles.description}">${escape(this.properties.description)}</p>
              <p class="${ styles.description}">${this.displayMode === DisplayMode.Edit ? "Edit Mode" : "Read Mode"}</p>
              <p class="${ styles.description}">Web Title: ${this.context.pageContext.web.title}</p>
              <p class="${ styles.description}">Email: ${this.context.pageContext.user.email}</p>
              <p class="${ styles.description}">${Environment.type === EnvironmentType.Local ? "LOCAL" : "SHAREPOINT"}</p>
              <p class="${ styles.description}">Continent where I reside: ${escape(this.properties.myContinent)}</p>
              <p class="${ styles.description}">Number of continents visited: ${this.properties.numContinentsVisited}</p>
              <a href="#" class="${ styles.button}">
                <span class="${ styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;

    Log.info('FirstSPFx', "Information Message", this.context.serviceScope);
    Log.warn('FirstSPFx', "Warning Message", this.context.serviceScope);
    Log.error('FirstSPFx', new Error("Error Message"), this.context.serviceScope);
    Log.verbose('FirstSPFx', "Verbose Message", this.context.serviceScope);

    this.domElement.getElementsByClassName(`${styles.button}`)[0]
      .addEventListener('click', (event: any) => {
        event.preventDefault();
        alert("Welcome to SPFX")
      })
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
                PropertyPaneTextField('myContinent', {
                  label: 'Continent where I currently reside',
                  onGetErrorMessage: this.validateContinents.bind(this)
                }),
                PropertyPaneSlider('numContinentsVisited', {
                  label: 'Number of continents visited',
                  min: 1,
                  max: 6,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private validateContinents(textValue: string): string {
    const options: string[] = ['africa', 'antarctica', 'asia', 'australia', 'north america', 'south america'];
    const inputToValidate: string = textValue.toLowerCase();

    return (options.indexOf(inputToValidate) === -1)?
    "Invalid continent entry"  : '';
  }
}
