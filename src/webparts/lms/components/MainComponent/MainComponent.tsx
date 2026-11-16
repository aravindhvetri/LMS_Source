import * as React from "react";
import Banner from "../Banner/Banner";
import "../../../../ExternalRef/CSS/Style.css";
import { graph } from "@pnp/graph";
const MainComponent = (props: any) => {
  const [isAdmin, setIsAdmin] = React.useState<any>(false);

  const getAdmin = async () => {
    const groupId = "c1a60f7a-a347-480e-9c6f-77e065517fb3"; 
    // Azure AD Group I  "bd693af8-d58a-48c8-8dc3-af8880eed9b4"

    await graph.groups
      .getById(groupId)
      .members()
      .then((members: any) => {
        const _filteredMembers = members.filter((member: any) => {
          return (
            member.userPrincipalName === props.context.pageContext.user.email
          );
        });
        if (_filteredMembers.length > 0) {
          setIsAdmin(true);
        } else {
          setIsAdmin(false);
        }
      })
      .catch((error: any) => {
        console.error("Error fetching group members:", error);
      });
  };

  React.useEffect(() => {
    void getAdmin();
  }, []);
  return (
    <Banner context={props.context} isAdmin={isAdmin} setIsAdmin={setIsAdmin} />
  );
};
export default MainComponent;
