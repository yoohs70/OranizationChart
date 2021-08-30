import React, { useState } from "react";
import { Dropdown, Image, Menu } from "@fluentui/react-northstar";
import "./Welcome.css";
import { EditCode } from "./EditCode";
import { AzureFunctions } from "./AzureFunctions";
import { Graph } from "./Graph";
import { CurrentUser } from "./CurrentUser";
import { useTeamsFx } from "./lib/useTeamsFx";
import { TeamsUserCredential } from "@microsoft/teamsfx";
import { useData } from "./lib/useData";
import { Admin } from "./Admin";
import { Publish } from "./Publish";
import { useGraph } from "./lib/useGraph";

import * as jquery from 'jquery';

// SharePoint
function getListItems(listName, reactHandler){
  var listItem;

  var deter = jquery.Deferred();
  
  // return jquery.ajax({
  //   url: ${reactHandler.props.siteurl}'/_api/web/lists/getbytitle('{listName}')/items',
  //   '
  // })
}


export function Welcome(props) {
  const { showFunction, environment } = {
    showFunction: true,
    environment: window.location.hostname === "localhost" ? "cart" : "admin",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      cart: "local environment",
      admin: "Azure environment",
    }[environment] || "local environment";

  // const steps = ["local", "azure", "publish"];
  const steps = ["cart", "admin", "publish"];
  const friendlyStepsName = {
    cart: "1. Organization Chart",
    admin: "2. Setting",
    publish: "3. Publish to Teams",
  };
  const [selectedMenuItem, setSelectedMenuItem] = useState("cart");
  const items = steps.map((step) => {
    return {
      key: step,
      content: friendlyStepsName[step] || "",
      onClick: () => setSelectedMenuItem(step),
    };
  });

  const { isInTeams } = useTeamsFx();
  const userProfile = useData(async () => {
    const credential = new TeamsUserCredential();
    return isInTeams ? await credential.getUserInfo() : undefined;
  })?.data;
  const userName = userProfile ? userProfile.displayName : "";
  
  // *** Department ***
  
  // will save department list items into below
  const [departmentItems, setDepartments] = useState();

  const { loading, error, data, reload } = useGraph(
    async (graph) => {
      // get all departments
      try {
        const resDepartment = await graph.api("/users")
         .select('Department')
         .get();

        // Remove duplicates
        const filteredDepartments = resDepartment.value.reduce((acc, current) => {
          const x = acc.find(item => item.department === current.department);
          if (!x) {
            return acc.concat([current]);
          } else {
            return acc;
          }
        }, []);

        setDepartments(filteredDepartments);

      } catch {
        // Could not fetch photo from user's profile, return empty string as placeholder.
        console.log("Error!!!!");
      }
      return { departmentItems };
    },
    { scope: ["User.Read.All", "User.ReadWrite.All", "Directory.Read.All", "Directory.ReadWrite.All"] }
    
  );

  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="hello.png" />
        <h1 className="center">Congratulations{userName ? ", " + userName : ""}!</h1>
        <p className="center">Your app is running in your {friendlyEnvironmentName}</p>
        <Menu defaultActiveIndex={0} items={items} underlined secondary />
        <div className="sections">
          {selectedMenuItem === "cart" && (
            <div>
              <EditCode showFunction={showFunction} />
              {isInTeams && <CurrentUser userName={userName} />}
              <Graph />
              {showFunction && <AzureFunctions />}

              {/* Departments  */}
              <h2>Departments</h2>
            
              <div class="dropdown">
                <select>
                  {
                    departmentItems && departmentItems.map((item) => <option>{item.department}</option>
                    )
                  }
                </select>
              </div>

              {
                departmentItems && departmentItems.map(
                  function(item, index){
                  return(
                  <tr>
                      <td>{item.department}</td>
                  </tr>);
                  }

                )
              }
            
            </div>

          )}

          
          {selectedMenuItem === "admin" && (
            <div>
              <Admin />
            </div>
          )}
          {selectedMenuItem === "publish" && (
            <div>
              <Publish />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
