import React, { useState } from "react";
import "./Admin.css";
import { Image } from "@fluentui/react-northstar";
import { useGraph } from "./lib/useGraph";

export function Admin(props) {

  // const { docsUrl } = {
  //   docsUrl: "https://aka.ms/teamsfx-docs",
  //   ...props,
  // };

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
    <div className="admin page">
      {/* <h2>Deploy to the Cloud</h2>
      <p>
        Before publishing your app to Teams App Catalog, you may want to provision and deploy your
        app's resources to the cloud to make sure your app will be running smoothly!
      </p>
      <p>
        To provision your resources, you can either use our CLI command "teamsfx provision" or apply
        “Teams: Provision in the Cloud" in Command palette.
      </p>
      <p>
        To deploy your app, you can either use our CLI command "teamsfx deploy" or apply “Teams:
        Deploy to the cloud" in Command palette.
      </p>
      <Image src="deploy.png" />
      <p>
        For more information, see the{" "}
        <a href={docsUrl} target="_blank" rel="noreferrer">
          docs
        </a>
        .
      </p> */}

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
  );
}
