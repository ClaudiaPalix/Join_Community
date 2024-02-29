import * as React from "react";
import { IJoinCommunityProps } from "./IJoinCommunityProps";
import { IJoinCommunityState } from "./IJoinCommunityState";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import styles from "./JoinCommunity.module.scss";
export default class JoinCommunity extends React.Component<
  IJoinCommunityProps,
  IJoinCommunityState
> {
  constructor(props: IJoinCommunityProps) {
    super(props);
    this.state = {
      allGroups: [],
      userTitle: [],
      allCommunities: [],
      allRequestCommunities: [],
    } as IJoinCommunityState;
    sp.setup({
      spfxContext: this.props.context,
    });
  }

  private company: string = "";

  // modifyUrl = () => {
  //   // Get the current URL
  //   let currentUrl = this.props.context.pageContext.web.absoluteUrl;

  //   // Replace the specified part of the URL
  //   currentUrl = currentUrl.replace(
  //     '/host/ad579878-7fab-46e5-be9a-2daa0487c87e/c80e60cd-8dd8-4192-aafe-46830fe88413',
  //     '/host/db5e5970-212f-477f-a3fc-2227dc7782bf/vivaengage'
  //   );
    
  //   console.log("Url Change:", currentUrl);
  //   // Change the URL
  //   window.location.href = currentUrl;
  // };

  private async _getCompanyFromEmail(): Promise<void> {

    console.log("User Email:", this.props.userEmail);
    if(this.props.userEmail.toLowerCase().includes("@aciesinnovations.com"))
    {
      console.log("User belongs to acies");
      this.company = "acies"
    }else if(this.props.userEmail.toLowerCase().includes("_zensar.com") || this.props.userEmail.includes("@zensar.")){
      console.log("User belongs to zensar");
      this.company = "zensar";
    }else if(this.props.userEmail.toLowerCase().includes("@rpg.com") || this.props.userEmail.includes("@rpg.in")){
      console.log("User belongs to rpg");
      this.company = "rpg";
    }else if(this.props.userEmail.toLowerCase().includes("@ceat.com")){
      console.log("User belongs to ceat");
      this.company = "ceat";
    }else if(this.props.userEmail.toLowerCase().includes("_harrisonsmalayalam.com") || this.props.userEmail.includes("@harrisonsmalayalam.com")){
      console.log("User belongs to harrison");
      this.company = "harrisonsmalayalam";
    }else if(this.props.userEmail.toLowerCase().includes("@kecrpg.com")){
      console.log("User belongs to kec");
      this.company = "kecrpg";
    }else if(this.props.userEmail.toLowerCase().includes("@raychemrpg.com")){
      console.log("User belongs to raychem");
      this.company = "raychemrpg";
    }else if(this.props.userEmail.toLowerCase().includes("@rpgls.com")){
      console.log("User belongs to rpgls");
      this.company = "rpgls";
    }
}  


  public async componentDidMount() {
    await this._getCompanyFromEmail();
    await this.fetchAllCommunity();
    await this.fetchAllRequestedCommunity();
    await this.getUserYammerGroups();
  }

  public render(): React.ReactElement<IJoinCommunityProps> {
    const decodedDescription = decodeURIComponent(this.props.description); // Decode the description (like incase there is blank space, or special characters, etc)
      console.log("Title: ",decodedDescription);
      const decodedSeeAllButton = decodeURIComponent(this.props.seeAllUrl);
          console.log("Url for See All button: ",decodedSeeAllButton);
          
        return (
          <div className={styles.joinCommunity} > 
            <div
              className="joinCommunity"
              style={{
                background: "#fff",
                width: "100%",
                padding: "20px",
                maxWidth:"500px",
              }}
            >
              <div className="topSection" style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                    marginBottom:"20px",
                  }}>
                <h2
                  style={{
                    color: "#20417c",
                    fontWeight: "500",
                    margin:"0",
                  }}
                >
                  {decodedDescription}
                </h2>
                <a href={decodedSeeAllButton} target="_self" data-interception="off" className={styles.SeeAllBtn}>See all</a>
              </div>
                <div>
                {this.state.allGroups.map((v: any) => {
                console.log("Group WebUrl:", v.CommunityURL.Url);
                let teamsLink = "";
                if (v.CommunityURL.Url && v.CommunityURL.Url.includes("groups/")) {
                    let splitUrl = v.CommunityURL.Url.split("groups/")[1];
                    console.log("SplitUrl:", splitUrl);
                    let teamsLink1 = "https://teams.microsoft.com/l/entity/db5e5970-212f-477f-a3fc-2227dc7782bf/vivaengage?context=%7B%22subEntityId%22:%22type=custom,data=group:";
                    let teamsLink2 = "%22%7D";
                    teamsLink = teamsLink1 + splitUrl + teamsLink2;
                    console.log("Full Link:", teamsLink);
                }
                
                // let isOutlook = false;
                let link = "";
                if(this.props.isTeams && this.props.isEmbedded){
                  link = teamsLink;
                // }else if(!this.props.isTeams && this.props.isEmbedded){
                //   link = "https://aka.ms/VivaEngage/Outlook";
                //   isOutlook = true
                }else{
                  link = v.CommunityURL.Url;
                }

                var metaTags = document.getElementsByTagName('meta');
                console.log("metaTags:", metaTags);

                // Loop through the meta tags to find the one with the name "publicUrl"
                for (var i = 0; i < metaTags.length; i++) {
                  var metaTag = metaTags[i];
                  
                  // Check if the meta tag has the name "publicUrl"
                  if (metaTag.getAttribute('name') === 'publicUrl') {
                    // Retrieve the content of the meta tag
                    var publicUrl = metaTag.getAttribute('content');
                    
                    // Log or use the publicUrl as needed
                    console.log('Public URL:', publicUrl);
                    
                    // Break out of the loop since we found the desired meta tag
                    break;
                  }
                }

                return (
                    <div className={styles.joinCommunityFieldBox}
                        key={v.field_1}
                        style={{
                            width: "100%",
                            display: "flex",
                            alignItems: "center",
                            flexWrap: "wrap",
                            justifyContent: "space-between",
                            background: "transparent",
                            border: "1px solid rgba(0, 0, 0, 0.1)",
                            padding: "11px 10px 10px",
                            marginTop: "10px",
                            font: "16px",
                            transition: ".3s ease-in-out",
                            boxSizing: "border-box",
                            fontWeight: "500",
                        }}
                    >
                        <a href={String(link)} target="_blank" style={{ marginBottom: "5px" }}>
                        {v.CommunityName}
                      </a>

                      {/* <a
                        href={String(link)}
                        target="_self"
                        style={{ marginBottom: "5px" }}
                        onClick={(event) => {
                          if (isOutlook) {
                            event.preventDefault(); // Prevent the default behavior of the click event
                            this.modifyUrl(); // Call this.modifyUrl if isOutlook is true
                          }
                          // Add any additional logic or fallback behavior if needed
                        }}
                      >
                        {v.CommunityName}
                      </a> */}
                        {this.isPresent(v)}
                    </div>
                );
            })}
                </div>
              </div>
            </div>
        );
      }

      isPresent(community: any) {
        let isCommunityPresent = false;
        this.state.allCommunities.forEach((e) => {
          if (e.full_name === community.CommunityName) {
            isCommunityPresent = true;
          }
        });
        if (isCommunityPresent) {
          return <strong>Already Joined</strong>;
        } else if (this.isRequested(community)) {
          return <strong>Already Requested</strong>;
        } else {
          return (
            <button  className={styles.joinCommunityBtn} 
              style={{
                backgroundColor: "#20417c",
                color: "#fff",
                alignItems: "center",
                padding: "6px 20px",
                border: "1px solid #20417c",
                cursor: "pointer",
                marginLeft: "5px",
                display: "block",
                transition: "0.3s ease-in-out",
                borderRadius:"3px",
                marginBottom:"5px",
              }}
              onClick={(e: any) => this.requestToJoin(community)}
            >
              Join
            </button>
          );
        }
      }
      private isRequested(community: any): boolean {
        let isCommunityPresent = false;

        this.state.allRequestCommunities.forEach((e) => {
          if (e.Title === community.CommunityName) {
            isCommunityPresent = true;
          }
        });
        return isCommunityPresent;
      }
      async requestToJoin(community: any) {
        //if (this.props.requestList !== undefined) {
        await sp.web.lists.getByTitle(`${this.props.selectedList2}`).items.add({
          Title: community.CommunityName,
          Community_x002f_GroupID: community.CommunityID,
          RequestorEmailID: this.props.userEmail,
          RequestApprovalStatus: "Pending",
        });
        alert("Successfully Requested");
        // } else {
        //   alert("Please fill the request list name from the property pane");
        // }
      }

      //
      async fetchAllCommunity() {
        // if (this.props.masterList !== undefined) {
          console.log('Selected List:', this.props.selectedList);
        const items2: any[] = await sp.web.lists
          // .getByTitle("Community_Master_List")
          .getByTitle(`${this.props.selectedList}`)
          .items.select("CommunityName", "CommunityID","Company", "CommunityURL")
          .orderBy("Order", true)
          .filter(`Company eq '${this.company.toUpperCase()}' or Company eq '${this.company.toLowerCase()}'`)();
          console.log('Fetched Communities:', items2);
        this.setState({
          allGroups: items2,
        });
        // }
      }
      async fetchAllRequestedCommunity() {
        // if (this.props.requestList !== undefined) {

        const items2: any[] = await sp.web.lists
          .getByTitle(`${this.props.selectedList2}`)
          .items.select("Title")
          .filter(`Status eq 'Pending'`)
          .filter(`RequestorEmailID eq '${this.props.userEmail}'`)
          .orderBy("Order", true)();
        this.setState({
          allRequestCommunities: items2,
        });

    //}
  }

  //
  async getUserYammerGroups() {
    await this.props.yammerProvider.getGroups().then((grps) => {
      this.setState({
        allCommunities: grps.data,
      });
    });
  }
}
