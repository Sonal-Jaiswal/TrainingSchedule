import * as React from "react";
import { PrimaryButton, DefaultButton } from "@fluentui/react";
import { IEnrollment, ITraining } from "../TrainingModels";

interface AdminBoardProps {
  trainings: ITraining[];
  enrollments: IEnrollment[];
  userRole: string;
  submitting: boolean;
  openTrainingForm: () => void;
  deleteTraining: (t: ITraining) => Promise<void>;
}

const AdminBoard: React.FC<AdminBoardProps> = ({ trainings, enrollments, userRole, submitting, openTrainingForm, deleteTraining }) => {
  return (
    <div style={{ padding: 0 }}>
      <h2 style={{ margin: 0, fontSize: 21, color: "#12343b", fontWeight: 600 }}>Admin Board</h2>
      <div style={{ display: "flex", gap: 10, margin: "12px 0 20px", flexWrap: "wrap" }}>
        {userRole === "Admin" && <PrimaryButton text="Create Training" onClick={openTrainingForm} />}
        {userRole === "Admin" && <DefaultButton text="Manage Training" onClick={() => {}} />}
        {userRole === "Admin" && <DefaultButton text="Manage Users" onClick={() => alert("User role management not configured") } />}
      </div>

      <h3 style={{ fontSize: 18, color: "#12343b", fontWeight: 600 }}>Training Management</h3>
      <div style={{ overflowX: "auto", marginBottom: 28 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ textAlign: "left", borderBottom: "2px solid #d9e2e7" }}>
              <th style={{ padding: 10 }}>Training</th>
              <th style={{ padding: 10 }}>Status</th>
              <th style={{ padding: 10 }}>Seats</th>
              <th style={{ padding: 10 }}>Action</th>
            </tr>
          </thead>
          <tbody>
            {trainings.map((training) => (
              <tr key={training.Id} style={{ borderBottom: "1px solid #edebe9" }}>
                <td style={{ padding: 10 }}>{training.TrainingName}</td>
                <td style={{ padding: 10 }}>{training.Status}</td>
                <td style={{ padding: 10 }}>{training.AvailableSeats}</td>
                <td style={{ padding: 10 }}>
                  <DefaultButton text="Delete" disabled={submitting} onClick={() => deleteTraining(training)} />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <h3 style={{ fontSize: 18, color: "#12343b", fontWeight: 600 }}>All Registrations</h3>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ textAlign: "left", borderBottom: "2px solid #d9e2e7" }}>
              <th style={{ padding: 10 }}>Employee</th>
              <th style={{ padding: 10 }}>Training</th>
              <th style={{ padding: 10 }}>Status</th>
              <th style={{ padding: 10 }}>Completion</th>
            </tr>
          </thead>
          <tbody>
            {enrollments.map((enrollment) => (
              <tr key={enrollment.Id} style={{ borderBottom: "1px solid #edebe9" }}>
                <td style={{ padding: 10 }}>{enrollment.Employee}</td>
                <td style={{ padding: 10 }}>{enrollment.Training}</td>
                <td style={{ padding: 10 }}>{enrollment.Status}</td>
                <td style={{ padding: 10 }}>{enrollment.CompletionStatus}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default AdminBoard;
