import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MkWebPart1WebPart.module.scss';
import * as strings from 'MkWebPart1WebPartStrings';
import welcomeDark from './assets/welcome-dark.png';
import welcomeLight from './assets/welcome-light.png';

export interface IMkWebPart1WebPartProps {
  description: string;
}

export default class MkWebPart1WebPart extends BaseClientSideWebPart<IMkWebPart1WebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.mkWebPart1}">
     <head><title style="color:blue;">IT Training and online service</title></head>
<body>
<marquee behavior="scroll">
<h3>🎓 Admissions Open 2026 | Free Demo Classes | 100% Placement Assistance | Enroll Today!</h3>
</marquee>
<table border="0" cellpadding="10" cellspacing="0" width="100%">
<tr> 
<th><p style="font-size:100px">&#128640;</p></th>
<th><h1 style="color:SkyBlue;">Transform Your Career with Industry-Focused Training</h1></th>
</tr>
</table>

<h2>Upcoming Batch Schedule</h2>
 
<table border="1" cellpadding="10" cellspacing="0" width="100%">
<tr style="background-color:#0078d4; color:white;">
<th>Course Name</th>
<th>Batch Start Date</th>
<th>Duration</th>
<th>Mode</th>
</tr>
 
<tr>
<td>Full Stack Development</td>
<td>05 November 2026</td>
<td>3 Months</td>
<td>Online / Classroom</td>
</tr>
 
<tr>
<td>Python Programming</td>
<td>15 November 2026</td>
<td>2 Months</td>
<td>Online</td>
</tr>
 
<tr>
<td>Microsoft Azure</td>
<td>01 December 2026</td>
<td>2 Months</td>
<td>Online / Classroom</td>
</tr>
 
<tr>
<td>DevOps with Azure</td>
<td>10 December 2026</td>
<td>3 Months</td>
<td>Online</td>
</tr>
 
<tr>
<td>Power BI</td>
<td>05 January 2027</td>

</body>
    </section>`;
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
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
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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
