import * as React from "react";
import { PrimaryButton } from "@fluentui/react";
import { ITraining } from "../TrainingModels";

interface TrainingCardProps {
  training: ITraining;
  userRole: string;
  openEnrollment: (t: ITraining) => void;
  styles: { [key: string]: React.CSSProperties };
}

const TrainingCard: React.FC<TrainingCardProps> = ({ training, userRole, openEnrollment, styles }) => {
  return (
    <div style={styles.trainingCard}>
      <div style={styles.trainingHeader}>
        <h3 style={styles.trainingTitle}>{training.TrainingName}</h3>
        <span style={training.Status.toLowerCase() === "active" ? styles.activeBadge : styles.inactiveBadge}>{training.Status}</span>
      </div>

      <p style={styles.description}>{training.Description}</p>

      <div style={styles.details}>
        <div style={styles.detailRow}><span style={styles.detailLabel}>Category</span><span style={styles.detailValue}>{training.Category}</span></div>
        <div style={styles.detailRow}><span style={styles.detailLabel}>Trainer</span><span style={styles.detailValue}>{training.Trainer}</span></div>
        <div style={styles.detailRow}><span style={styles.detailLabel}>Training Date</span><span style={styles.detailValue}>{training.TrainingDate ? new Date(training.TrainingDate).toLocaleDateString() : "Not specified"}</span></div>
        <div style={styles.detailRow}><span style={styles.detailLabel}>Available Seats</span><span style={{ ...styles.detailValue, fontWeight: 600 }}>{training.AvailableSeats}</span></div>
      </div>

      {userRole === "Employee" && (
        <PrimaryButton text={training.AvailableSeats > 0 ? "Enroll Now" : "Fully Booked"} disabled={training.AvailableSeats <= 0} onClick={() => openEnrollment(training)} styles={{ root: { width: "100%", borderRadius: "6px" } }} />
      )}
    </div>
  );
};

export default TrainingCard;
