import * as React from 'react';
import styles from './ArcherSerHwFrmWp.module.scss';
import { IArcherSerHwFrmWpProps } from './IArcherSerHwFrmWpProps';
import { IArcherSerHwFrmWpState } from './IArcherSerHwFrmWpState';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton ,mergeStyles, PrimaryButton, Label, Stack, MessageBar, MessageBarType, Separator } from 'office-ui-fabric-react';
import { sp, DateTimeFieldFormatType } from "@pnp/sp/presets/all";
import { FormatDate } from '../../Utilities';
import { ISiteUserProps } from "@pnp/sp/site-users/";


export default class ArcherSerHwFrmWp extends React.Component<IArcherSerHwFrmWpProps, IArcherSerHwFrmWpState> {
  constructor(props: IArcherSerHwFrmWpProps, state: IArcherSerHwFrmWpState) {
    super(props);
    const url = new URL(window.location.href);
    const params = new URLSearchParams(url.search);
    
    let qsParam: string;
    params.has('idval') ? qsParam = params.get("idval") : qsParam = "";
    
    this.state = {
      itemID: qsParam,
      Site: "",
      IP: "",
      HostName: "",
      Function: "",
      OperationStatus: "",
      DNAAlias: "",
      OS: "",
      SLALevel: "",
      BackupPolicy:[],
      SignOffStatus:"",
      Version:"",
      Created: "",
      CreatedBy: "",
      Modified: null,
      ModifiedBy: "",
    };
    sp.setup({
      spfxContext: this.props.spcontext
    });
    
    this._getItem(Number(this.state.itemID));
  }
  private  _closeClicked(): void {
    
    window.history.back();
  }
  private async _getItem(qid:number) {
    // get a specific item by id
    const item: any = await sp.web.lists.getByTitle("Server Hardware")     
      .items.getById(qid) 
      .select("*","OData__UIVersionString")
      .get();


      
    let uservalue: number = item["AuthorId"];
    let DisplayUserCreated:string;
        try
        { 
            const user: ISiteUserProps = await sp.web.getUserById(uservalue).get();
            DisplayUserCreated= user.Title;
            
        }
        catch(error){  
          DisplayUserCreated="User no longer exist in our systems ";
         
        }  
      let editorvalue: number = item["EditorId"];
        let DisplayUserEdited:string;
            try
            { 
                const user: ISiteUserProps = await sp.web.getUserById(editorvalue).get();
                DisplayUserEdited= user.Title;
                
            }
            catch(error){   
              DisplayUserEdited="User Deleted from systems";
            
            }  
    
  

    console.dir(item);  
//set value 

    this.setState({
      HostName: item.Title,
      itemID: item.itemID,
      Site: item.Site,
      IP: item.IP,
      Function: item.Function,
      OperationStatus: item.Operation_x0020_Status,
      DNAAlias: item.DNS_x0020_Alias,
      OS: item.OS,
      SLALevel: item.SLA_x0020_Level,
      BackupPolicy:item.Backup_x0020_Policy[0],
      SignOffStatus:"",
      Version:item.OData__UIVersionString,
      Created: FormatDate(item.Created),
      CreatedBy: DisplayUserCreated,
      Modified: FormatDate(item.Modified),
      ModifiedBy: DisplayUserEdited,

    });
  }

  

    public render(): React.ReactElement<IArcherSerHwFrmWpProps> {
   return (
      <div className={styles.archerSerHwFrmWp}>     
     
     
      <div className={styles.mystyles}>
          <span className={styles.btnalignright}>
           <PrimaryButton  text="Back" onClick={this._closeClicked} />
        </span>
        <span><h2>Server Hardware Details</h2></span>
        <div className={styles.mytablestyles}>
        <table >
          <tr>
            <td className={styles.valTdColspan}>
              <span> <Label className={styles.mylabel}>HostName :</Label></span>
            </td>
            <td>
              <span> <Label className={styles.valLabel}>{this.state.HostName}</Label></span>
            </td>
          </tr>
          <tr>
            <td>
              <span> <Label className={styles.mylabel}>DNA Alias :</Label></span>
            </td>
            <td>
              <span> <Label className={styles.valLabel}>{this.state.DNAAlias}</Label></span>
            </td>
          </tr>
        
          <tr>
            <td>
              <span> <Label className={styles.mylabel}>Site:</Label></span>
            </td>
            <td>
              <span> <Label className={styles.valLabel}>{this.state.Site}</Label></span>
            </td>
          </tr>
          <tr>
            <td>
              <span> <Label className={styles.mylabel}>IP :</Label></span>
            </td>
            <td>
              <span> <Label className={styles.valLabel}>{this.state.IP}</Label></span>
            </td>
          </tr>
          
          <tr>
            <td>
              <span> <Label className={styles.mylabel}>Function :</Label></span>
            </td>
            <td>
              <span> <Label className={styles.valLabel}>{this.state.Function}</Label></span>
            </td>
          </tr>
        
          <tr>
            <td>
              <span> <Label className={styles.mylabel}>Operation Status : </Label></span>
            </td>
            <td>
              <span> <Label className={styles.valLabel}>{this.state.OperationStatus}</Label></span>
            </td>
          </tr>
          <tr>
            <td>
              <span> <Label className={styles.mylabel}>OS :</Label></span>
            </td>
            <td>
              <span> <Label className={styles.valLabel}>{this.state.OS}</Label></span>
            </td>
          </tr>
          <tr>
            <td>
              <span> <Label className={styles.mylabel}>SLA Level :</Label></span>
            </td>
            <td>
              <span> <Label className={styles.valLabel}>{this.state.SLALevel}</Label></span>
            </td>
          </tr>

          <tr>
            <td>
              <span> <Label className={styles.mylabel}>Backup Policy :</Label></span>
            </td>
            <td>
              <span> <Label className={styles.valLabel}>{this.state.BackupPolicy}</Label></span>
            </td>
          </tr>         
          <tr>
            <td>
              <span> <Label className={styles.mylabel} disabled>Version :</Label></span>
            </td>
            <td>
              <span> <Label  className={styles.valLabel} disabled>{this.state.Version}</Label></span>
            </td>
          </tr>
          <tr>
            <td>
              <span> <Label  className={styles.mylabel} disabled>Created :</Label></span>
            </td>
            <td>
              <span> <Label  className={styles.valLabel} disabled>{this.state.Created} - {this.state.CreatedBy}</Label></span>
            </td>
          </tr>

          <tr>
            <td>
              <span> <Label  className={styles.mylabel} disabled>Last Modifed :</Label></span>
            </td>
            <td>
              <span> <Label  className={styles.valLabel} disabled>{this.state.Modified} - {this.state.ModifiedBy}</Label></span>
            </td>
          </tr>

        </table>
      </div>
      </div>
   </div>
      
    );
  }
}
