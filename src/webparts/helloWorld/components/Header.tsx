import * as React from "react";
import { DefaultButton, PrimaryButton } from "@fluentui/react";

interface IHeaderProps {
  userRole: string;
  employeeName: string;
  employeeInitials: string;
  showProfile: boolean;
  setShowProfile: (v: boolean) => void;
  showMyCourses: boolean;
  toggleMyCourses: () => void;
  showAdminBoard: boolean;
  toggleAdminBoard: () => void;
  onOpenTrainingForm: () => void;
  setError: (msg: string) => void;
}

const Header: React.FC<IHeaderProps> = ({
  userRole,
  employeeName,
  employeeInitials,
  showProfile,
  setShowProfile,
  showMyCourses,
  toggleMyCourses,
  showAdminBoard,
  toggleAdminBoard,
  onOpenTrainingForm,
  setError
}) => {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
      <div>
        <h1 style={{ margin: "0 0 8px 0", fontSize: 28, fontWeight: 600 }}>{userRole} Training Dashboard</h1>
        <p style={{ margin: 0, fontSize: 15, opacity: 0.9 }}>Discover and enroll in internal training programs as a {userRole.toLowerCase()}</p>
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

        {userRole !== "Admin" && (
          <DefaultButton text={showMyCourses ? "All Trainings" : "My Courses"} onClick={toggleMyCourses} styles={{ root: { borderRadius: "6px" } }} />
        )}

        {userRole === "Admin" && (
          <>
            <DefaultButton text={showAdminBoard ? "Training Catalog" : "Admin Board"} onClick={() => { toggleAdminBoard(); }} styles={{ root: { borderRadius: "6px" } }} />
            <PrimaryButton text="Create Training" onClick={onOpenTrainingForm} />
          </>
        )}
      </div>
    </div>
  );
};

export default Header;
