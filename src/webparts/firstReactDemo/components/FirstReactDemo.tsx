import * as React from 'react';
import styles from './FirstReactDemo.module.scss';
import { IFirstReactDemoProps } from './IFirstReactDemoProps';
import { IFirstReactDemoState } from './IFirstReactDemoState';
import { escape } from '@microsoft/sp-lodash-subset';
import SPOperations from '../../services/SPServices';
import {DefaultButton, Dropdown, IDropdownOption} from 'office-ui-fabric-react';

export default class FirstReactDemo extends React.Component<IFirstReactDemoProps, IFirstReactDemoState> {

  private _spOps: SPOperations;
  private selectedListTile:string;

  public constructor(props:IFirstReactDemoProps){
    super(props);
    this._spOps = new SPOperations();
    this.state = {
      listTitles:[],
      status:""
    };
  }

  public getListTitle = (event:any, data:any) =>{
    this.selectedListTile = data.text;
  }

  public componentDidMount(){
    this._spOps.GetAllListsTitles(this.props.context).then((results:IDropdownOption[]) => {
        this.setState({listTitles: results});
    });
  }


  public render(): React.ReactElement<IFirstReactDemoProps> {

    return (
      <div className={ styles.firstReactDemo }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SPFX React Demo</span>
              <p className={ styles.subTitle }>DEMO: CRUD Operations using Rest API (spHTTPClient</p>
                <div id="divParent" className={styles.myStyles}>
                  <Dropdown
                  className={styles.dropdown}
                  options={this.state.listTitles}
                  placeholder="--- Select your list ---"
                  onChange={this.getListTitle}
                  >

                  </Dropdown>
                  <DefaultButton
                  className={styles.myButton}
                  text="Create List Item"
                  onClick={() =>
                    this._spOps
                    .CreateListItem(this.props.context,this.selectedListTile)
                    .then((result:string) =>{
                    this.setState({status: result});
                  })}
                  ></DefaultButton>

                  <div>{this.state.status}</div>
            </div>
            </div>

          </div>
        </div>
      </div>
    );
  }
}
