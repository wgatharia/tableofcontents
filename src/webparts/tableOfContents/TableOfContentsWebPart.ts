import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TableOfContentsWebPart.module.scss';
import * as strings from 'TableOfContentsWebPartStrings';

import { spfi, SPFx } from '@pnp/sp';

import "@pnp/sp/webs";
import "@pnp/sp/search";
import "@pnp/sp/hubsites";
import "@pnp/sp/hubsites/web";

import { IHubSiteWebData, IHubSiteInfo } from  "@pnp/sp/hubsites";

import { ISearchQuery, SearchQueryBuilder } from "@pnp/sp/search";
import { forEach } from 'lodash';

export interface ITableOfContentsWebPartProps {
  description: string;
  tocSource: string;
  sortBy: string;
  sortOrder: number;
  rowLimit: number;
}

export default class TableOfContentsWebPart extends BaseClientSideWebPart<ITableOfContentsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public async render(): Promise<void> {
    const sp = spfi().using(SPFx(this.context));
    
    const webData: Partial<IHubSiteWebData> = await sp.web.hubSiteData();
    const siteId = this.context.pageContext.site.id.toString();
    const siteTitle = this.context.pageContext.web.title;

    let DepartmentId = webData.parentHubSiteId; 

    const hubsites: IHubSiteInfo[] = await sp.hubSites();

    const isHubSite = hubsites.filter(x => x.SiteId == siteId).length > 0;

    let html: string = '';

    if(isHubSite) {
      DepartmentId = siteId;
    }
    // The source of Sites to List is either a Hub or List Subwebs on a Site.
    // For a Hub use Search to filter by Department Id ( Hub Site Id ).
    if(this.properties.tocSource === 'Hub') {
      let queryTemplate = `{searchterms} ((contentclass=STS_Site OR contentclass=STS_Web) AND NOT (IsHubSite:true)) (DepartmentId:${DepartmentId} OR DepartmentId:{${DepartmentId}})`;

      console.log(queryTemplate);
  
      const appSearchSettings: ISearchQuery = {
        QueryTemplate: queryTemplate,
        ClientType: "Custom",
        RowLimit: this.properties.rowLimit,
        RowsPerPage:this.properties.rowLimit,        
        SelectProperties: ["ContentType","ContentTypeId","Title","SiteName","SiteTitle","SPWebUrl", "WebPath","PreviewUrl","IconUrl","ClassName","LastModifiedTime"],
        Properties: [],
        SortList: [{Property: `${this.properties.sortBy}`, Direction: this.properties.sortOrder}],
        TrimDuplicates:false
      };
  
      const builder = SearchQueryBuilder("", appSearchSettings);
  
      const results = await sp.search(builder);
      console.log(results.RawSearchResults.PrimaryQueryResult.RelevantResults.Table.Rows);
      if(results.TotalRows > 0) {
        let rows = results.RawSearchResults.PrimaryQueryResult.RelevantResults.Table.Rows;
        
        rows.forEach((row) => {
          let link = row.Cells.filter((x) => x.Key == "SPWebUrl")[0].Value;
          let title = row.Cells.filter((x) => x.Key == "SiteTitle")[0].Value;
          let icon = link + "/_api/siteiconmanager/getsitelogo?type='1'";

          html += `<div class=${styles.listItem}><a href=${link}><div class=${styles.listItem}><img src=${icon} class=${styles.bannerImage} alt="." /><span class=${styles.titleSpan}>${title}</span></div></a></div>`;
        });

        this.domElement.innerHTML = `
        <section class="${styles.tableOfContents} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
          <div>
            <p class=${styles.welcome}>${siteTitle} Sites</p>
          </div>
          <div>
            ${html}
          </div>
        </section>`;    
      } else {
        this.domElement.innerHTML = `
        <section class="${styles.tableOfContents} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
          <div>
            <p class=${styles.welcome}>There are no Sites to display. Ensure to add this web part to a Hub Site and there are sites associated with the Hub!</p>
          </div>
        </section>`;
      }
    }

    if(this.properties.tocSource === 'Site') {
      const subWebs = await sp.web.getSubwebsFilteredForCurrentUser().select("Title", "ServerRelativeUrl").orderBy("Title", true)();
      if(subWebs.length > 0) {
          html += `<div>`;
          subWebs.forEach(async (web) => {
            let icon = web.ServerRelativeUrl + "/_api/siteiconmanager/getsitelogo?type='1'";
            html += `<div class=${styles.listItem}><img src=${icon} class=${styles.bannerImage} /><a href=${web.ServerRelativeUrl}><span class=${styles.titleSpan}>${web.Title}</span></a></div>`;
          });
          html += `</div>`;
          this.domElement.innerHTML = `
          <section class="${styles.tableOfContents} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
            <div>
              <p class=${styles.welcome}>${siteTitle} Subsites</p>
            </div>
            <div>
              ${html}
            </div>
          </section>`; 
      } else {
        this.domElement.innerHTML = `
        <section class="${styles.tableOfContents} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
          <div>
            <p class=${styles.welcome}>There are no Sub Sites to display! Try editing and changing web part Toc Source.</p>
          </div>
        </section>`;        
      }
    }
    this.renderCompleted();
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
        return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
                PropertyPaneChoiceGroup('tocSource', {
                  label: strings.TocSourceFieldLabel,
                  options: [
                    { key: "Site", text: "This Site"},
                    { key: "Hub", text: "All Sites in the Hub", checked: true},
                  ]
                }),
                PropertyPaneDropdown("sortBy", {
                  label: strings.SortByFieldLabel,
                  disabled: this.properties.tocSource === "Site",
                  options: [
                    { key: "SiteName", text: "Site Name"},
                    { key: "LastModifiedTime", text: "Last modified"}
                  ],
                  selectedKey: "SiteName"
                }),
                PropertyPaneDropdown("sortOrder",{
                  label: strings.SortOrderLabel,
                  disabled: this.properties.tocSource === "Site",
                  options: [
                    { key: 0, text: "Ascending"},
                    { key: 1, text: "Descending"}
                  ],
                  selectedKey: 0
                }),
                PropertyPaneSlider("rowLimit", {
                  label: strings.RowLimitLabel,
                  min: 10,
                  max: 50,
                  value: 10,
                  showValue: true,
                  step: 5
                }) 
              ]
            }
          ]
        }
      ]
    };
  }
}
