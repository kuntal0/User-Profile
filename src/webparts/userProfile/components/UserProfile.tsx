import * as React from "react";
// import styles from './UserProfile.module.scss';
import { IUserProfileProps } from "./IUserProfileProps";
// import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from "@microsoft/sp-http";
//import styles from "./UserProfile.module.scss";
// import styles from "./UserProfile.module.scss";

interface Profile {
  givenName: string;
  jobTitle: string;
  mail: string;
  mobilePhone: string;
  officeLocation: string;
  preferredLanguage: string;
  surname: string;
  displayName: string;
}

export default class UserProfile extends React.Component<
  IUserProfileProps,
  Profile
> {
  constructor(props: IUserProfileProps, state: Profile) {
    super(props);
    this.state = {
      givenName: "",
      jobTitle: "",
      mail: "",
      mobilePhone: "",
      officeLocation: "",
      preferredLanguage: "",
      surname: "",
      displayName: "",
    };
  }

  componentDidMount(): void {
    this.getmyProfile();
  }

  public getmyProfile() {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("/me")
          .version("v1.0")
          // .select("displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage:,surname")
          .get((err: any, res: any) => {
            this.setState({
              displayName: res.displayName,
              mail: res.mail,
              jobTitle: res.jobTitle,
              givenName: res.givenName,
              mobilePhone: res.mobilePhone,
              officeLocation: res.officeLocation,
              preferredLanguage: res.preferredLanguage,
              surname: res.surname,
            });
            console.log(res);
          });
      });
  }

  public render(): React.ReactElement<IUserProfileProps> {
    // const {
    //   description,
    //   isDarkTheme,
    //   environmentMessage,
    //   hasTeamsContext,
    //   userDisplayName
    // } = this.props;

    return (

      <section>
        <div>
          <h1>
            <b>{this.state.displayName}</b>
          </h1>
          <p>Given Name : {this.state.givenName},</p>
          <p>Surname : {this.state.surname},</p>
          <p>Mail ID : {this.state.mail},</p>
        </div>
        <img src="https://graph.microsoft.com/v1.0/me/photo/$value" alt="" />
      </section>
    );
  }
}