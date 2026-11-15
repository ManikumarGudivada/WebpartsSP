import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpBirthdayWebpart1WebPart.module.scss';
import * as strings from 'SpBirthdayWebpart1WebPartStrings';
import welcomeDark from './assets/welcome-dark.png';
import welcomeLight from './assets/welcome-light.png';

export interface ISpBirthdayWebpart1WebPartProps {
  // description: string;
  Id: number;
  Title: string;
  Birthday: string;
  profile?: any;
}

interface IBirthdayEmployee {
  Id: number;
  Title: string;
  Birthday: string;
  profile?: any;
}
export default class SpBirthdayWebpart1WebPart extends BaseClientSideWebPart<ISpBirthdayWebpart1WebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  
  private readonly listViewUrl: string =
    'https://ford728.sharepoint.com/sites/PoliceDocumentation/Lists/Harika12/AllItems.aspx';

  public render(): void {
    this.loadBirthdays().catch((error: unknown) => {
      console.error('Birthday loading error:', error);
    });
  }

  private async loadBirthdays(): Promise<void> {

    this.domElement.innerHTML = `
      <div class="${styles.birthdayContainer}">
        <h2>🎂 Team Birthday Wishes</h2>
        <p>Loading birthdays...</p>
      </div>
    `;

    try {

      const listUrl = new URL(this.listViewUrl);
      const serverRelativeListUrl = listUrl.pathname.replace(
        /\/AllItems\.aspx\/?$/i,
        ''
      );
      const listsPathIndex = serverRelativeListUrl.toLowerCase().indexOf('/lists/');

      if (listsPathIndex < 0) {
        throw new Error('The configured URL must point to a SharePoint list view.');
      }

      const listWebUrl =
        `${listUrl.origin}${serverRelativeListUrl.substring(0, listsPathIndex)}`;
      const escapedListUrl = serverRelativeListUrl.replace(/'/g, "''");

      const apiUrl =
        `${listWebUrl}/_api/web/GetList('${escapedListUrl}')/items?$select=Title,Birthday,profile&$top=5000`;

      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          apiUrl,
          SPHttpClient.configurations.v1
        );

      if (!response.ok) {
        const responseDetails = await response.text();
        throw new Error(
          `SharePoint error: ${response.status} ${response.statusText}. ${responseDetails}`
        );
      }

      const data = await response.json();

      const employees: IBirthdayEmployee[] = data.value;

      this.displayBirthdays(employees);

    } catch (error) {

      console.error('Birthday loading error:', error);

      this.domElement.innerHTML = `
        <div class="${styles.error}">
          <h3>Unable to load birthdays</h3>
          <p>Please check your SharePoint List and column names.</p>
          <p>${this.escapeHtml(String(error))}</p>
        </div>
      `;
    }
  }

  private escapeHtml(value: string): string {

    return value.replace(/[&<>"']/g, (char: string) => {

      const entities: { [key: string]: string } = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
      };

      return entities[char];
    });
  }

  private getPhotoUrl(profile: any): string {

    if (!profile) {
      return '';
    }

    let photoUrl = '';

    if (typeof profile === 'string') {

      photoUrl = profile;

    } else if (typeof profile === 'object') {

      if (profile.serverRelativeUrl) {

        photoUrl = profile.serverUrl
          ? profile.serverUrl + profile.serverRelativeUrl
          : profile.serverRelativeUrl;

      } else if (profile.Url) {

        photoUrl = profile.Url;

      } else if (profile.url) {

        photoUrl = profile.url;
      }
    }

    if (!photoUrl) {
      return '';
    }

    try {

      const parsedUrl = new URL(
        photoUrl,
        this.context.pageContext.web.absoluteUrl
      );

      if (
        parsedUrl.protocol !== 'https:' &&
        parsedUrl.protocol !== 'http:'
      ) {
        return '';
      }

      return parsedUrl.href;

    } catch {

      return '';
    }
  }

  private getBirthdayDays(birthday: string): number {

    const dob = new Date(birthday);

    if (isNaN(dob.getTime())) {
      return 999;
    }

    const today = new Date();

    const todayDate = new Date(
      today.getFullYear(),
      today.getMonth(),
      today.getDate()
    );

    const birthMonth = dob.getUTCMonth();
    const birthDate = dob.getUTCDate();

    let nextBirthday = new Date(
      today.getFullYear(),
      birthMonth,
      birthDate
    );

    if (nextBirthday.getTime() < todayDate.getTime()) {

      nextBirthday = new Date(
        today.getFullYear() + 1,
        birthMonth,
        birthDate
      );
    }

    return Math.round(
      (nextBirthday.getTime() - todayDate.getTime()) /
      (1000 * 60 * 60 * 24)
    );
  }

  private formatBirthday(birthday: string): string {

    const dob = new Date(birthday);

    if (isNaN(dob.getTime())) {
      return 'Birthday unavailable';
    }

    return dob.toLocaleDateString('en-IN', {
      day: 'numeric',
      month: 'long',
      timeZone: 'UTC'
    });
  }

  private displayBirthdays(
    employees: IBirthdayEmployee[]
  ): void {

    const todayEmployees: IBirthdayEmployee[] = [];
    const upcomingEmployees: IBirthdayEmployee[] = [];

    employees.forEach((employee) => {

      if (!employee.Birthday) {
        return;
      }

      const days = this.getBirthdayDays(employee.Birthday);

      if (days === 0) {

        todayEmployees.push(employee);

      } else if (days <= 30) {

        upcomingEmployees.push(employee);
      }
    });

    upcomingEmployees.sort((a, b) => {

      return this.getBirthdayDays(a.Birthday) -
        this.getBirthdayDays(b.Birthday);
    });

    const createCard = (
      employee: IBirthdayEmployee,
      isToday: boolean
    ): string => {

      const name = this.escapeHtml(
        employee.Title || 'Team Member'
      );

      const birthday = this.formatBirthday(
        employee.Birthday
      );

      const photoUrl = this.getPhotoUrl(
        employee.profile
      );

      const photo = photoUrl
        ? `
          <img
            class="${styles.employeePhoto}"
            src="${this.escapeHtml(photoUrl)}"
            alt="${name}"
          />
        `
        : `
          <div class="${styles.defaultPhoto}">
            👤
          </div>
        `;

      const wish = isToday
        ? `
          <div class="${styles.todayWish}">
            🎂 Happy Birthday!
          </div>
          <p class="${styles.message}">
            Wishing you happiness, success and a wonderful year ahead!
          </p>
        `
        : `
          <div class="${styles.upcomingWish}">
            🎉 Birthday in ${this.getBirthdayDays(employee.Birthday)} days
          </div>
        `;

      return `
        <div class="${styles.birthdayCard}">

          ${photo}

          <h3>${name}</h3>

          <p class="${styles.date}">
            ${birthday}
          </p>

          ${wish}

        </div>
      `;
    };

    const todayHtml = todayEmployees.length > 0
      ? todayEmployees.map((employee) =>
          createCard(employee, true)
        ).join('')
      : `
        <div class="${styles.noBirthday}">
          <span>🌸</span>
          <p>No birthdays today!</p>
        </div>
      `;

    const upcomingHtml = upcomingEmployees.length > 0
      ? upcomingEmployees.map((employee) =>
          createCard(employee, false)
        ).join('')
      : `
        <div class="${styles.noBirthday}">
          <p>No upcoming birthdays in the next 30 days.</p>
        </div>
      `;

    this.domElement.innerHTML = `

      <div class="${styles.birthdayContainer}">

        <div class="${styles.header}">

          <div class="${styles.confetti}">
            🎈 🎉 🎂 🎉 🎈
          </div>

          <h1>Team Birthday Wishes</h1>

          <p>Celebrating our amazing team members!</p>

        </div>

        <section class="${styles.birthdaySection}">

          <h2>🎂 Today's Birthdays</h2>

          <div class="${styles.cardGrid}">
            ${todayHtml}
          </div>

        </section>

        <section class="${styles.birthdaySection}">

          <h2>📅 Upcoming Birthdays</h2>

          <div class="${styles.cardGrid}">
            ${upcomingHtml}
          </div>

        </section>

        <div class="${styles.footer}">
          Made with ❤️ for our team
        </div>

      </div>
    `;
  }

  protected get dataVersion(): Version {

    return Version.parse('1.0');
  }
  // public render(): void {
  //   this.domElement.innerHTML = `
  //   <section class="${styles.spBirthdayWebpart1}">
  //     <div class="${styles.welcome}">
  //       <img alt="" src="${this._isDarkTheme ? welcomeDark : welcomeLight}" class="${styles.welcomeImage}" />
  //       <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
  //       <div>${this._environmentMessage}</div>
  //       <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
  //     </div>
  //     <div>
  //       <h3>Welcome to SharePoint Framework!</h3>
  //       <p>
  //       The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
  //       </p>
  //       <h4>Learn more about SPFx development:</h4>
  //         <ul class="${styles.links}">
  //           <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
  //           <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
  //         </ul>
  //     </div>
  //   </section>`;
  // }


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

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

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
