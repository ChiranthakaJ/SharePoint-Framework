import * as React from 'react';
import Tile from "./Tile";
import "./FirstReactJSWebPart.css";
import type { IFirstReactJsWebPartProps } from './IFirstReactJsWebPartProps';

export default class FirstReactJsWebPart extends React.Component<IFirstReactJsWebPartProps> {
  public render(): React.ReactElement<IFirstReactJsWebPartProps> {

    // Define the base URL for the images stored in the SharePoint Document Library
    const siteUrl = "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/Site%20Images/";


    // Sample tiles array with title, icon, and link properties
    const tiles = [
      { title: "Projects", icon: "Projects.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/PROJECTS/" },
      { title: "Collateral", icon: "Collateral.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/COLLATERAL/" },
      { title: "Corporate", icon: "Corporate.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/CORPORATE/" },
      { title: "Governance", icon: "Governance.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/GOVERNANCE%20%20HR/" },
      { title: "Events", icon: "Events.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/EVENTS/" },
      { title: "Visitor Economy", icon: "Visitor%20Economy.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/VISITOR%20ECONOMY/" },
      { title: "Communications", icon: "Communications.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/COMMUNICATIONS/" },
      { title: "Finance & HR", icon: "Finance%20and%20HR.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/FINANCE%20%20FUNDING/" },
      { title: "Image Library", icon: "Image%20Library.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/IMAGE%20LIBRARY/" },
      { title: "Board", icon: "Board.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/BOARD/" },
      { title: "Staff Resources", icon: "Staff%20Resources.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/RESOURCES/" },
      { title: "Archive", icon: "Archive.png", link: "https://centralhighlandsrc.sharepoint.com/sites/CHDCHub/ARCHIVE/" },
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
