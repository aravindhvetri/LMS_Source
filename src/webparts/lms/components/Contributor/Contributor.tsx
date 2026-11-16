import * as React from "react";
import styles from "./Contributor.module.scss";
import { Carousel } from "primereact/carousel";
import { sp } from "@pnp/sp/presets/all";
import { Label } from "@fluentui/react";

const Contributor = (props: any) => {
  const [topContributors, setTopContributors] = React.useState<any[]>([]);

  const getTopContributors = () => {
    const countByEmail: Record<string, { count: number; lastDate: string }> = {};

    // Count contributions & track last event date
    props.EventData.forEach((event: any) => {
      const email = event.CreatedByEmail;
      if (!email) return;

      const eventDate = new Date(event.Date);
      if (!countByEmail[email]) {
        countByEmail[email] = { count: 0, lastDate: event.Date };
      }

      countByEmail[email].count++;
      if (new Date(countByEmail[email].lastDate) < eventDate) {
        countByEmail[email].lastDate = event.Date;
      }
    });

    // Fetch profile details
    Promise.all(
      Object.entries(countByEmail).map(([email, { count, lastDate }]) => {
        let role = "";
        let name = email;
        let imageUrl = "";

        return sp.web
          .ensureUser(email)
          .then((ensuredUser: any) => {
            const loginName = ensuredUser.data.LoginName;

            const foundUser = props.SiteUsers.find(
              (u: any) => u.secondaryText === email
            );

            if (foundUser) {
              name = foundUser.text || email;
              imageUrl = foundUser.imgUrl || "";
            }

            return sp.profiles
              .getUserProfilePropertyFor(loginName, "Title")
              .then((userRole: string) => {
                role = userRole;
                return {
                  name,
                  email,
                  role,
                  count,
                  lastEventDate: lastDate,
                  imageUrl,
                };
              });
          })
          .catch(() => {
            // handle if user profile or ensureUser fails
            return { name, email, role, count, lastEventDate: lastDate, imageUrl };
          });
      })
    )
      .then((contributorArray) => {
        const top3 = contributorArray
          .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name))
          .slice(0, 3)
          .map((contributor, index) => ({ ...contributor, rank: index + 1 }));

        setTopContributors(top3);
      })
      .catch((error) => {
        console.error("Error fetching top contributors:", error);
        setTopContributors([]);
      });
  };

  // 👉 Custom contributor message
  const getCustomMessage = (contributor: any) => {
    switch (contributor.rank) {
      case 1:
        return `${contributor.name.split(" ")[0]} has consistently gone above and beyond, contributing ${contributor.count} times! A true leader and inspiration to the team.`;
      case 2:
        return `${contributor.name.split(" ")[0]} is a dedicated team player who has made ${contributor.count} strong contributions, always supporting our initiatives.`;
      case 3:
        return `${contributor.name.split(" ")[0]} shows great enthusiasm and commitment with ${contributor.count} valuable contributions so far. Keep growing!`;
      default:
        return `${contributor.name.split(" ")[0]} is an active contributor.`;
    }
  };

  const contributorTemplate = (contributor: any) => (
    <div className={styles.contributor_card}>
      <img
        src={contributor.imageUrl}
        alt={contributor.name}
        className={styles.profile_img}
      />
      <h3 className={styles.name}>{contributor.name}</h3>

      <div
        className={`${styles.badge} ${
          contributor.rank === 1
            ? styles.gold
            : contributor.rank === 2
            ? styles.silver
            : styles.bronze
        }`}
      >
        {contributor.rank === 1
          ? "🥇 #1 Top Contributor"
          : contributor.rank === 2
          ? "🥈 #2 Contributor"
          : "🥉 #3 Contributor"}
      </div>

      <p className={styles.description}>
        {getCustomMessage(contributor)} The most recent contribution was on{" "}
        <b style={{ color: "#747171" }}>
          {new Date(contributor.lastEventDate).toLocaleDateString("en-GB")}
        </b>
        .
      </p>
    </div>
  );

  React.useEffect(() => {
    getTopContributors();
  }, []);

  return (
    <>
      <h4 className={styles.Title}>Contributor Spotlight</h4>
      <div className={styles.ContributorLeader}>
        {topContributors.length > 0 ? (
          <Carousel
            value={topContributors}
            numVisible={1}
            numScroll={1}
            circular
            autoplayInterval={2500}
            className={styles.custom_carousel}
            itemTemplate={contributorTemplate}
          />
        ) : (
          <div className={styles.nodatafound}>
            <Label>No contributor found ...</Label>
          </div>
        )}
      </div>
    </>
  );
};

export default Contributor;
