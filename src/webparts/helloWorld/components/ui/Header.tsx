import * as React from "react";
import { DefaultButton } from "@fluentui/react";

interface HeaderProps {
  userRole: string;
  employeeName: string;
  employeeInitials: string;
  showProfile: boolean;
  setShowProfile: (v: boolean) => void;
  showMyCourses: boolean;
  setShowMyCourses: (v: boolean) => void;
  showAdminBoard: boolean;
  setShowAdminBoard: (v: boolean) => void;
}

const Header: React.FC<HeaderProps> = ({
  userRole,
  employeeName,
  employeeInitials,
  showProfile,
  setShowProfile,
  showMyCourses,
  setShowMyCourses,
  showAdminBoard,
  setShowAdminBoard
}) => {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
      <div>
        <h1 style={{ margin: 0, fontSize: "28px", fontWeight: 600 }}>{userRole} Training Dashboard</h1>
        <p style={{ margin: 0, fontSize: "15px", opacity: 0.9 }}>
          Discover and enroll in internal training programs as a {userRole.toLowerCase()}
        </p>
      </div>

      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 10, position: "relative" }}>
        <button
          type="button"
          aria-label="Open profile"
          title={employeeName || "Open profile"}
          onClick={() => setShowProfile(!showProfile)}
          style={{ width: 52, height: 52, borderRadius: "50%", border: "2px solid rgba(255,255,255,0.85)", backgroundColor: "#ffffff", color: "#007f86", fontSize: 17, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 2px 8px rgba(0,0,0,0.18)" }}
        >
          {employeeInitials}
        </button>
        {showProfile && (
          <div style={{ position: "absolute", top: 64, right: 0, width: 210, padding: 14, backgroundColor: "#ffffff", color: "#12343b", borderRadius: 8, boxShadow: "0 8px 20px rgba(0,0,0,0.2)", zIndex: 10, textAlign: "left" }}>
            <p style={{ margin: 0, fontSize: 15, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{employeeName || "User"}</p>
            <p style={{ margin: "5px 0 0", color: "#52666d", fontSize: 12 }}>Role: {userRole}</p>
          </div>
        )}

        {userRole !== "Admin" && (
          <DefaultButton
            text={showMyCourses ? "All Trainings" : "My Courses"}
            onClick={() => setShowMyCourses(!showMyCourses)}
            styles={{ root: { borderRadius: "6px" }, rootHovered: { borderRadius: "6px" } }}
          />
        )}

        {userRole === "Admin" && (
          <DefaultButton
            text={showAdminBoard ? "Training Catalog" : "Admin Board"}
            onClick={() => {
              setShowAdminBoard(!showAdminBoard);
              setShowMyCourses(false);
            }}
            styles={{ root: { borderRadius: "6px" }, rootHovered: { borderRadius: "6px" } }}
          />
        )}
      </div>
    </div>
  );
};

export default Header;
