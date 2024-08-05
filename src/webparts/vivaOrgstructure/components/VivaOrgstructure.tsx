import * as React from 'react';
import { SPFI } from '@pnp/sp';
import { getSP } from '../pnpjsConfig';
import styles from './VivaOrgstructure.module.scss';
import classNames from 'classnames';
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
        .items.select('OrgStructureJSON, FullName, AccountName, Position')
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
          .items.select('OrgStructureJSON, FullName, AccountName, Position')
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
    this.setState({ loginName: e.target.id }, () => { this.getOrgStructure() });
  }


  private requestPhotoUrl(userLogin: string) {
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
                <div className={styles.userInfo} id={userItem.Manager.Login}><p>{userItem.Manager.Name} {userItem.Manager.FamilyName}</p><span>{userItem.Manager.Position}</span></div>
              </div>)}
          </div>) : (<div>Немає інформації</div>)}
        {this.state.orgUser.length > 0 ? (
          <div className={styles.vivaOrgstructureUser}>
            {this.state.orgUser.map((userItem: any) =>
              <div className={styles.container}>
                <div className={styles.userPhoto}>
                  <img src={this.requestPhotoUrl(userItem.Login)} draggable="false" />
                </div>
                <div className={styles.userInfo}><p>{userItem.FullName}</p><span></span>{userItem.Position}</div>
              </div>)}
          </div>) : (<div>Немає інформації</div>)}
        {this.state.orgEmployees.length > 0 ? (
          <div className={styles.vivaOrgstructureEmployees}>
            {this.state.orgEmployees.map((employeeItem: any) =>
              <div className={styles.container} onClick={(e: any) => this.onClickItem(e)}>
                <div className={styles.userPhoto}>
                  <img src={this.requestPhotoUrl(employeeItem.AccountName)} id={employeeItem.AccountName} draggable="false" />
                </div>
                <div className={styles.userInfo} id={employeeItem.AccountName}><p>{employeeItem.FullName}</p><span></span>{employeeItem.Position}</div>
              </div>)}
          </div>) : (<div>Немає інформації</div>)}
        <div className={classNames(styles.ms_PersonaCard)}>
          <div className={classNames(styles.ms_PersonaCard_persona)}>
            <div className={classNames(styles.ms_Persona, styles.ms_Persona__lg)}>
              <div className={classNames(styles.ms_Persona_imageArea)}>
                <div className={classNames(styles.ms_Persona_initials, styles.ms_Persona_initials__blue)}>AL</div>
                <img className={classNames(styles.ms_Persona_image)} src={this.requestPhotoUrl(this.state.loginName)} />
              </div>
              <div className={classNames(styles.ms_Persona_details)}>
                <div className={classNames(styles.ms_Persona_primaryText)}>Панюта Олександр Олександрович</div>
                <div className={classNames(styles.ms_Persona_secondaryText)}>Розробник ПЗ на базі SharePiont</div>
              </div>
            </div>
          </div>
          <ul className={classNames(styles.ms_PersonaCard_actions)}>
            <li data-action-id="chat" className={classNames(styles.ms_PersonaCard_action)}>
              <i className="ms-Icon ms-Icon--Chat"></i>
            </li>
            <li data-action-id="phone" className={classNames(styles.ms_PersonaCard_action, styles.is_active)} >
              <i className="ms-Icon ms-Icon--Phone"></i>
            </li>
            <li data-action-id="video" className={classNames(styles.ms_PersonaCard_action)}>
              <i className="ms-Icon ms-Icon--Video"></i>
            </li>
            <li data-action-id="mail" className={classNames(styles.ms_PersonaCard_action)} >
              <i className="ms-Icon ms-Icon--Mail"></i>
            </li>
          </ul>
          <div className={classNames(styles.ms_PersonaCard_actionDetailBox)}>
            <div data-detail-id="mail" className={classNames(styles.ms_PersonaCard_details)}>
              <div className={classNames(styles.ms_PersonaCard_detailLine)}><span className={classNames(styles.ms_PersonaCard_detailLabel)}>Персональний E-mail: </span>
                <a className={classNames(styles.ms_Link)} href="mailto: alton.lafferty@outlook.com">alton.lafferty@outlook.com</a>
              </div>
              <div className={classNames(styles.ms_PersonaCard_detailLine)}><span className={classNames(styles.ms_PersonaCard_detailLabel)}>Робочий E-mail: </span>
                <a className={classNames(styles.ms_Link)} href="mailto: alton.lafferty@outlook.com">altonlafferty@contoso.com</a>
              </div>
            </div>
            <div data-detail-id="phone" className={classNames(styles.ms_PersonaCard_details)}>
              <div className={classNames(styles.ms_PersonaCard_detailLine)}>
                <span className={classNames(styles.ms_PersonaCard_detailLabel)}>Додаткова інформація</span>
              </div>
              <div className={classNames(styles.ms_PersonaCard_detailLine)}>
                <span className={classNames(styles.ms_PersonaCard_detailLabel)}>Персональний телефон:</span>  555.206.2443
              </div>
              <div className={classNames(styles.ms_PersonaCard_detailLine)}>
                <span className={classNames(styles.ms_PersonaCard_detailLabel)}>Робочий телефон:</span>  555.929.8240
              </div>
            </div>
          </div>
        </div>
        AL

      </div >
    );
  }
}
