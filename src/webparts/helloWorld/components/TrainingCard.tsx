import * as React from "react";
import { PrimaryButton } from "@fluentui/react";
import { ITraining } from "./TrainingModels";

interface IProps {
  training: ITraining;
  userRole: string;
  onEnroll: (t: ITraining) => void;
}

const TrainingCard: React.FC<IProps> = ({ training, userRole, onEnroll }) => {
  return (
    <div style={{ backgroundColor: "#ffffff", border: "1px solid #d9e2e7", borderRadius: 12, padding: 20, boxShadow: "0 2px 8px rgba(0,0,0,0.06)" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, marginBottom: 12 }}>
        <h3 style={{ margin: 0, color: "#12343b", fontSize: 19, fontWeight: 600 }}>{training.TrainingName}</h3>
        <span style={training.Status.toLowerCase() === "active" ? { backgroundColor: "#dff6dd", color: "#107c10", padding: "5px 10px", borderRadius: 15, fontSize: 12, fontWeight: 600 } : { backgroundColor: "#fde7e9", color: "#a4262c", padding: "5px 10px", borderRadius: 15, fontSize: 12, fontWeight: 600 }}>{training.Status}</span>
      </div>
      <p style={{ color: "#52666d", fontSize: 14, lineHeight: 1.5, minHeight: 63, marginBottom: 16 }}>{training.Description}</p>
      <div style={{ borderTop: "1px solid #d9e2e7", paddingTop: 14, marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 10, padding: "6px 0", fontSize: 13 }}>
          <span style={{ color: "#52666d", fontWeight: 600 }}>Category</span>
          <span style={{ color: "#12343b", textAlign: "right" }}>{training.Category}</span>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 10, padding: "6px 0", fontSize: 13 }}>
          <span style={{ color: "#52666d", fontWeight: 600 }}>Trainer</span>
          <span style={{ color: "#12343b", textAlign: "right" }}>{training.Trainer}</span>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 10, padding: "6px 0", fontSize: 13 }}>
          <span style={{ color: "#52666d", fontWeight: 600 }}>Training Date</span>
          <span style={{ color: "#12343b", textAlign: "right" }}>{training.TrainingDate ? new Date(training.TrainingDate).toLocaleDateString() : "Not specified"}</span>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 10, padding: "6px 0", fontSize: 13 }}>
          <span style={{ color: "#52666d", fontWeight: 600 }}>Available Seats</span>
          <span style={{ color: "#12343b", textAlign: "right", fontWeight: 600 }}>{training.AvailableSeats}</span>
        </div>
      </div>

      {userRole === "Employee" && (
        <PrimaryButton text={training.AvailableSeats > 0 ? "Enroll Now" : "Fully Booked"} disabled={training.AvailableSeats <= 0} onClick={() => onEnroll(training)} styles={{ root: { width: "100%", borderRadius: "6px" } }} />
      )}
    </div>
  );
};

export default TrainingCard;
