import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AnnoucementWebPart.module.scss';
import * as strings from 'AnnoucementWebPartStrings';
import { getIconClassName } from '@uifabric/styling';



import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import { Fields } from '@pnp/sp/fields';
import { ICamlQuery } from "@pnp/sp/lists";
import "@pnp/sp/content-types";
import { ContentTypes, IContentTypes, IContentType, IContentTypeInfo } from "@pnp/sp/content-types";

export interface IAnnoucementWebPartProps {
  listName: string;
}

export interface IAnnouncements {
  value: IAnnouncement[];
}
export interface IAnnouncement {
  Id: string;
  Title: string;
  Body: string;
}

export default class AnnoucementWebPart extends BaseClientSideWebPart<IAnnoucementWebPartProps> {
  private disableListNameTextBox: boolean = true;
  private customAnnouncementCTID: string = "";
  private listAnnouncementCTID: string = "";

  public render(): void {
    this.domElement.innerHTML = `<div class="">  
                                  <header class="${styles.mnHeader}">  
                                      <h4>Meldinger</h4>         
                                    </header>  
                                    <div id="announcementListContainer">  
                                    </div>  
                                  </div>  
                                  </div>`;

    if (this.properties.listName !== "" &&
      this.properties.listName !== undefined &&
      this.disableListNameTextBox) {
      const caml: ICamlQuery =
      {
        ViewXml: `<View>
                                                  <RowLimit>5</RowLimit>
                                              
                                                  <Query>
                                                  <Where>
                                                     <Eq>
                                                        <FieldRef Name='ContentType' />
                                                        <Value Type='Computed'>Announcement-VTFK</Value>
                                                     </Eq>
                                                  </Where>
                                                  <OrderBy>
                                                     <FieldRef Name='Created' Ascending='False' />
                                                  </OrderBy>
                                               </Query>
                                              </View>
                                              `,
      };

      const listData = sp.web.lists.getByTitle(this.properties.listName).getItemsByCAMLQuery(caml);

      listData.then((result) => {
        if (this.listAnnouncementCTID === "") {
          const associatedCT = sp.web.lists.getByTitle(this.properties.listName).contentTypes;
          var lCtIT = "";

          associatedCT.get().then((data) => {
            data.forEach(function (item, index) {
              if (item.Name === "Announcement-VTFK")//Announcement-VTFK
              {
                lCtIT = item.StringId;
                console.log("Fetched CT: " + lCtIT);

                return true;
              }
            });
          }).then(() => {
            this.RenderHtmlFromData(result, lCtIT);
          }
          );
        }
        else {
          this.RenderHtmlFromData(result, this.listAnnouncementCTID);
        }
      });
    }
  }

  private RenderHtmlFromData(announcements: any[], lCtIT: string): void {
    this.listAnnouncementCTID = lCtIT;
    sp.web.lists.getByTitle(this.properties.listName).select('DefaultViewUrl').get().then((dt) => {
      console.log(dt.DefaultViewUrl);
      var lstRelUrl = dt.DefaultViewUrl.substr(0, dt.DefaultViewUrl.lastIndexOf("/"));
      let html: string = '';
      var listRefUrl = this.GetDominUrl() + lstRelUrl;

      announcements.forEach((item) => {

        var dom = document.createElement('div');
        dom.innerHTML = item.Body;
        var bText = dom.textContent;

        if (bText !== "" && bText.length > 30) {
          bText = bText.substring(0, 99) + "...";
        }

        var viewUrl = listRefUrl + "/DispForm.aspx?ID=" + item.ID

        html += `<div class="${styles.media}">
                <div class="${styles.icHolder}"><i class="${getIconClassName('CampaignTemplate')}"></i></div>
                <div class="${styles.mediabody}">
                  <h5 class=""><a href="${viewUrl}" target="_blank">${item.Title}</a></h5>
                  <p>${bText}</p>
                </div>
              </div>`
      });

      html += `<div>  
                  <a href="${listRefUrl}/NewForm.aspx?ContentTypeId=${this.listAnnouncementCTID}"  target="_blank" class="${styles.btnCustom}"><i class="${getIconClassName('AddTo')}"> </i> Legg til ny kunngjøring</a>  
                </div>     
              `;

      const listContainer: Element = this.domElement.querySelector('#announcementListContainer');
      listContainer.innerHTML = html;
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this.properties.listName !== undefined &&
      this.properties.listName !== "") {
      this.disableListNameTextBox = true;
      this.context.propertyPane.refresh();
    }
    else {
      this.disableListNameTextBox = false;
    }

    this.context.propertyPane.refresh();
  }

  //Called when web part is initialised 
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // other init code may be present
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  private ButtonClickProvisionList(oldVal: any): any {
    if (this.properties.listName !== undefined && this.properties.listName !== "") {
      const curSiteContentTypes = sp.web.contentTypes.select('StringId', 'Name').get();
      var isCTExists = false;
      var cID = "";

      curSiteContentTypes.then((data) => {
        if (data.length !== undefined && data.length > 0) {

          data.forEach(function (item, index) {
            if (item.Name === "Announcement-VTFK")//Announcement-VTFK 
            {
              isCTExists = true;
              cID = item.StringId;
              return true;
            }
          });
        }
      }).then((d)=>
      {
        if (isCTExists) {
          this.customAnnouncementCTID = cID;
          this.ProvisionList();
        }
        else {
          console.log("Content Type not found");
        }
      });
    }
  }

  private GetDominUrl()
  {
    var arr = window.location.href.split("/");
    var domainUrl = arr[0] + "//" + arr[2];

    return domainUrl;
  }

  private ProvisionList() {
    if (this.customAnnouncementCTID !== "") {
      //const web = Web(this.context.pageContext.web.absoluteUrl);
      const lst = sp.web.lists.ensure(this.properties.listName, "Announcement List", 100);

      lst.then((listExistResult) => {
        if (listExistResult.created) {
          sp.web.lists.getByTitle(this.properties.listName).contentTypes.addAvailableContentType(this.customAnnouncementCTID);
          this.disableListNameTextBox = true;
        }
        else {
          var confRes = confirm("List with name " + this.properties.listName +
            " already exists. If you want to use this list click OK, else click Cancel" +
            " to create another list with different name.");
          if (confRes) 
          {
            this.disableListNameTextBox = true;
            const associatedCT = sp.web.lists.getByTitle(this.properties.listName).contentTypes;
            var alreadyCTAssociated = false;

            associatedCT.get().then((data) => {
              data.forEach(function (item, index) {
                if (item.Name === "Announcement-VTFK")//Announcement-VTFK
                {
                  alreadyCTAssociated = true;
                  return true;
                }
              });
            }).then((d)=>{
              if (!alreadyCTAssociated) {
                sp.web.lists.getByTitle(this.properties.listName).contentTypes.addAvailableContentType(this.customAnnouncementCTID);
                console.log("Content type associated with list successfully");
              }
              else
              {
                console.log("Content type already associated with list");
              }
            });
          }
          else {
            this.disableListNameTextBox = false;
          }
        }

        //Once provision completes refresh content
        this.context.propertyPane.refresh();
        this.render();
      });
    }
    else {
      console.log("There are some issue while accessing ContentType");
    }
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
                PropertyPaneTextField('listName', {
                  label: "Announcement List Name:",
                  disabled: this.disableListNameTextBox
                }),
                PropertyPaneButton('addList',
                  {
                    text: "",
                    buttonType: PropertyPaneButtonType.Hero,
                    onClick: this.ButtonClickProvisionList.bind(this),
                    icon: 'Add',
                    disabled: this.disableListNameTextBox
                  })
              ]
            }
          ]
        }
      ]
    };
  }
}
