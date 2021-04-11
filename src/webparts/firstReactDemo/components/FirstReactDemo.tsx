import * as React from 'react';
import styles from './FirstReactDemo.module.scss';
import { IFirstReactDemoProps } from './IFirstReactDemoProps';
import { IFirstReactDemoState } from './IFirstReactDemoState';
import { escape } from '@microsoft/sp-lodash-subset';
import SPOperations from '../../services/SPServices';
import {Dropdown, IDropdownOption} from 'office-ui-fabric-react';

export default class FirstReactDemo extends React.Component<IFirstReactDemoProps, IFirstReactDemoState> {

  private _spOps: SPOperations;

  public constructor(props:IFirstReactDemoProps){
    super(props);
    this._spOps = new SPOperations();
    this.state = {
      listTitles:[]
    };
  }

  public componentDidMount(){
    this._spOps.GetAllList(this.props.context).then((results:IDropdownOption[]) => {
        this.setState({listTitles: results});
    });
  }


  public render(): React.ReactElement<IFirstReactDemoProps> {

    let options:IDropdownOption[] = [];

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
                  placeholder="--- Select your list ---">
                  </Dropdown>
            </div>
            </div>

          </div>
        </div>
      </div>
    );
  }
}
