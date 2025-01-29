import * as React from 'react';
import Tile from "./Tile";
import "./FirstReactJSWebPart.css";
import type { IFirstReactJsWebPartProps } from './IFirstReactJsWebPartProps';

export default class FirstReactJsWebPart extends React.Component<IFirstReactJsWebPartProps> {
  public render(): React.ReactElement<IFirstReactJsWebPartProps> {

    // Define the base URL for the images stored in the SharePoint Document Library
    const siteUrl = "https://fourthar.sharepoint.com/sites/My-Developer-Communication-Site/Site%20Images/";


    // Sample tiles array with title, icon, and link properties
    const tiles = [
      { title: "Projects", icon: "Projects.png", link: "https://www.microsoft.com" },
      { title: "Collateral", icon: "Collateral.png", link: "https://learn.microsoft.com" },
      { title: "Corporate", icon: "Corporate.png", link: "https://outlook.office.com" },
      { title: "Governance", icon: "Governance.png", link: "https://teams.microsoft.com" },
      { title: "Events", icon: "Events.png", link: "https://outlook.office.com/calendar" },
      { title: "Visitor Economy", icon: "Visitor%20Economy.png", link: "https://onedrive.live.com" },
      { title: "Communications", icon: "Communications.png", link: "https://yourcompany.sharepoint.com" },
      { title: "Finance & HR", icon: "Finance%20and%20HR.png", link: "#" },
      { title: "Image Library", icon: "Image%20Library.png", link: "#" },
      { title: "Board", icon: "Board.png", link: "#" },
      { title: "Staff Resources", icon: "Staff%20Resources.png", link: "#" },
      { title: "Archive", icon: "Archive.png", link: "#" },
    ];

    

    return (
      <section>
      <div className="parent-container"> {/* Full-width wrapper */}
        <div className="tile-container">
          {tiles.map((tile, index) => (
            <Tile
              key={index}
              title={tile.title}
              icon={`${siteUrl}${tile.icon}`}  // Construct the correct URL here
              link={tile.link}
            />
          ))}
        </div>
      </div>
    </section>
    );
  }
}
