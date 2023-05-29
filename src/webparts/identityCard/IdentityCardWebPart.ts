import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './IdentityCardWebPart.module.scss';
import * as strings from 'IdentityCardWebPartStrings';

import * as moment from 'moment';

export interface IIdentityCardWebPartProps {
  cognome: string;
  nome: string;
  luogoDiNascita: string;
  genere: string;
  dataDiNascita: string;
  immagineBase64: string;
}

export default class IdentityCardWebPart extends BaseClientSideWebPart<IIdentityCardWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    //lint trick
    this._environmentMessage;

    this.domElement.innerHTML = `
    <section class="${styles.identityCard} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.card} ${this._isDarkTheme ? 'dark':''}">
        <div class="${styles.data}">
          <ul>
            <li>
              <span class="${styles.bold}">${escape(this.properties.cognome)} ${escape(this.properties.nome)}</span>
            </li>
            <hr/>
            <li>
              <span class="${styles.bold}">${strings.LuogoDiNascitaFieldLabel}:</span> <span>${escape(this.properties.luogoDiNascita)}</span>
            </li>
            <li>
              <span class="${styles.bold}">${strings.GenereFieldLabel}:</span> <span>${escape(this.properties.genere)}</span>
            </li>
            <li>
              <span class="${styles.bold}">${strings.DataDiNascitaFieldLabel}:</span> <span>${escape(this.properties.dataDiNascita)}</span>
            </li>
          </ul>
        </div>
        <div class="${styles.picture}">
          <img src="data:image/jpeg;base64,${escape(this.properties.immagineBase64)}">
        </div>
      </div>
    </section>`;
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      this.properties.cognome = "Di Trani";
      this.properties.nome = "Nicola";
      this.properties.luogoDiNascita = "Andria";
      this.properties.dataDiNascita = "11/11/1986";
      this.properties.genere = "Male";
      this.properties.immagineBase64 = "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";
    });
  }

  private validateDescription(value: string): string {
    const dateFormat = 'DD/MM/YYYY';
    if ((!!value && value.length != 10) || !moment(value, dateFormat).isValid()) {
      return strings.DateFormatErrorMessage;
    }
    return '';
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.DettaglioUtentePaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('cognome', {
                  label: strings.CognomeFieldLabel
                }),
                PropertyPaneTextField('nome', {
                  label: strings.NomeFieldLabel
                }),
                PropertyPaneTextField('luogoDiNascita', {
                  label: strings.LuogoDiNascitaFieldLabel
                }),
                PropertyPaneTextField('genere', {
                  label: strings.GenereFieldLabel
                }),
                PropertyPaneTextField('dataDiNascita', {
                  label: strings.DataDiNascitaFieldLabel,
                  validateOnFocusIn:true,
                  validateOnFocusOut:true,
                  onGetErrorMessage: this.validateDescription.bind(this)
                }),
                PropertyPaneTextField('immagineBase64', {
                  label: strings.ImmagineBase64FieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
