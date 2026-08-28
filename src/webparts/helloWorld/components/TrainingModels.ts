// Represents one training record from the Trainings-SAR SharePoint list.
export interface ITraining {
 Id: number;
 TrainingName: string;
 Description: string;
 Category: string;
 Trainer: string;
 TrainingDate: string;
 AvailableSeats: number;
 Status: string;
}

// Represents one enrollment record from the Enrollments-SAR SharePoint list.
export interface IEnrollment {
 Id: number;
 Employee: string;
 Training: string;
 TrainingId: number;
 EnrollmentDate: string;
 Status: string;
 CompletionStatus: string;
}
