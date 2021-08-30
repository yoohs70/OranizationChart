import React from "react";
//import { Avatar, Card, Flex, Skeleton, Text } from "@fluentui/react-northstar";

export const DepartmentsCard = ( data) => (
    <div>
        <h2>Departments</h2>
        { data.departmentItems && data.departmentItems.map(
            function(item, index){
            return(
            <tr>
                <td>{item.department}</td>
            </tr>);
            }

        )}
        
    </div>

);