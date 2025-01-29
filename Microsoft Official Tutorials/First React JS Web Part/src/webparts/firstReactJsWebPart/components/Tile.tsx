import * as React from "react";
import "./Tile.css";

interface TileProps {
  title: string;
  icon: string;
  link?: string;
}

const Tile: React.FC<TileProps> = ({ title, icon, link }) => {
    return (
      <a href={link || "#"} className="tile" target="_blank" rel="noopener noreferrer" style={{ backgroundImage: `url(${icon})` }}>
        <span className="tile-title">{title}</span>
      </a>
    );
  };

export default Tile;
