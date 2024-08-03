import * as React from 'react';
import { SPFI } from '@pnp/sp';
import { getSP } from '../pnpjsConfig';
import styles from './VivaOrgstructure.module.scss';
import type { IVivaOrgstructureProps, IVivaOrgstructureState } from './IVivaOrgstructure';

export default class VivaOrgstructure extends React.Component<IVivaOrgstructureProps, IVivaOrgstructureState> {

  private _sp: SPFI;

  constructor(props: IVivaOrgstructureProps) {
    super(props);
    this.state = {
      orgUser: {},
      orgEmployees: [],
      loginName: this.props.startLogin
    };
    this._sp = getSP(this.props.context);
  }

  componentDidMount() {
    this.getOrgStructure();
  }

  private async getOrgStructure(): Promise<void> {
    //Get user info
    console.log('state', this.state);
    try {
      const userItem: any = await this._sp.web
        .getList(this.props.context.pageContext.web.serverRelativeUrl + this.props.employeeListLink)
        .items.select('OrgStructureJSON, FullName, AccountName')
        .filter(`AccountName eq '${this.state.loginName}'`)
        .top(1)();
      console.log('userItem', userItem);
      if (userItem[0].OrgStructureJSON !== null) {
        let userItemJSON: any = JSON.parse(userItem[0].OrgStructureJSON);
        userItemJSON.FullName = (userItem[0].FullName) === null ? '' : userItem[0].FullName;
        var userItemArray: any = Array(userItemJSON);
        console.log(userItemArray);
        //Get employee info
        var employeeItem: any = await this._sp.web
          .getList(this.props.context.pageContext.web.serverRelativeUrl + this.props.employeeListLink)
          .items.select('OrgStructureJSON, FullName, AccountName')
          .filter(`ManagerAccount eq '${userItemJSON.Manager.Login}'`)
          .top(1000)();
        console.log('employeeItem', employeeItem);
      }
      this.setState({ orgUser: (userItemArray) ? userItemArray : [], orgEmployees: (employeeItem) ? employeeItem : [] });
    } catch (error) {
      console.log('Помилка формування оргструктури!', error);
    }
  }

  private onClickItem(e: any) {
    console.log('EventID', e.target.id);
    this.setState({ loginName: e.target.id });
    this.getOrgStructure();
  }


  private requestPhotoUrl(userLogin: string) {
    console.log('this.props.context.pageContext', this.props.context.pageContext);
    return `${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?UserName=${userLogin}@${this.props.loginDomain}&size=S`;
  }

  render() {
    return (
      <div className={styles.vivaOrgstructureWrapper}>
        {this.state.orgUser.length > 0 ? (
          <div className={styles.vivaOrgstructureManager} >
            {this.state.orgUser.map((userItem: any) =>
              <div className={styles.container} onClick={(e: any) =>
                this.onClickItem(e)}>
                <div className={styles.userPhoto}>
                  <img src={this.requestPhotoUrl(userItem.Manager.Login)} id={userItem.Manager.Login} draggable="false" />
                </div>
                <div className={styles.userInfo} id={userItem.Manager.Login}>{userItem.Manager.Name} {userItem.Manager.FamilyName}</div>
              </div>)}
          </div>) : (<div>Немає інформації</div>)}
        {this.state.orgUser.length > 0 ? (
          <div className={styles.vivaOrgstructureUser}>
            {this.state.orgUser.map((userItem: any) =>
              <div className={styles.container}>
                <div className={styles.userPhoto}>
                  <img src={this.requestPhotoUrl(userItem.Login)} draggable="false" />
                </div>
                <div className={styles.userInfo}>{userItem.FullName}</div>
              </div>)}
          </div>) : (<div>Немає інформації</div>)}
        {this.state.orgEmployees.length > 0 ? (
          <div className={styles.vivaOrgstructureEmployees}>
            {this.state.orgEmployees.map((employeeItem: any) =>
              <div className={styles.container} onClick={(e: any) => this.onClickItem(e)}>
                <div className={styles.userPhoto}>
                  <img src={this.requestPhotoUrl(employeeItem.AccountName)} id={employeeItem.AccountName} draggable="false" />
                </div>
                <div className={styles.userInfo} id={employeeItem.AccountName}>{employeeItem.FullName}</div>
              </div>)}
          </div>) : (<div>Немає інформації</div>)}
      </div>
    );
  }
}
