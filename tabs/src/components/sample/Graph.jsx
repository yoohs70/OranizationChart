import React, {
  useState
} from 'react';

import { Button } from "@fluentui/react-northstar";
import { useGraph } from "./lib/useGraph";
import { ProfileCard } from "./ProfileCard";
import { DepartmentsCard } from "./DepartmentsCard";

export function Graph() {
/*
  const { loading, error, data, reload } = useGraph(
    async (graph) => {
      const profile = await graph.api("/me").get();
      let photoUrl = "";
      try {
        const photo = await graph.api("/me/photo/$value").get();
        photoUrl = URL.createObjectURL(photo);
      } catch {
        // Could not fetch photo from user's profile, return empty string as placeholder.
      }
      return { profile, photoUrl };
    },
    { scope: ["User.Read"] }
  );
*/
  const [departmentItems, setDepartments] = useState();

  const { loading, error, data, reload } = useGraph(
   
   async (graph) => {    
    const profile = await graph.api("/me").get(); 
    let photoUrl = "";
    try {
      const photo = await graph.api("/me/photo/$value").get();
      photoUrl = URL.createObjectURL(photo);
    } catch {
      // Could not fetch photo from user's profile, return empty string as placeholder.
    }

    try {
       const resDepartment = await graph.api("/users")
        .select('Department')
        .get();
       setDepartments(resDepartment.value);
     } catch {
       // Could not fetch photo from user's profile, return empty string as placeholder.
       console.log("Error!!!!");
     }
     return { profile, departmentItems };
   },
   { scope: ["User.Read", "User.Read.All", "User.ReadWrite.All", "Directory.Read.All", "Directory.ReadWrite.All"] }
 );
   

  return (
    <div>
      <h2>Get the user's profile photo</h2>
      <p>Click below to authorize this app to read your profile photo using Microsoft Graph.</p>
      <Button primary content="Authorize" disabled={loading} onClick={reload} />
      {/* {loading && ProfileCard(true)} */}

      {/* <div>
        <h2>Departments</h2>
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
      </div> */}


      {loading && ProfileCard(true)}
      {!loading && error && <div className="error">{error.toString()}</div>}
      {!loading && data && ProfileCard(false, data)}
    </div>
  );
}
