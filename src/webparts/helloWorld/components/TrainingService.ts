import { SPFI } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";

import {
 ITraining,
 IUserRole,
 UserRole
} from "./TrainingModels";

export class UnauthorizedError extends Error {
 public constructor(message: string = "You are not authorized to perform this action.") {
   super(message);
   this.name = "UnauthorizedError";
 }
}

export const canManageTrainings = (role: UserRole): boolean =>
 role === "Admin" || role === "HR";

export const canDeleteTrainings = (role: UserRole): boolean =>
 role === "Admin";

export const canViewAllEnrollments = (role: UserRole): boolean =>
 role === "Admin" || role === "HR";

export class TrainingService {
 private readonly sp: SPFI;

 public constructor(sp: SPFI) {
   this.sp = sp;
 }

 public async getCurrentUserRole(): Promise<UserRole> {
   const currentUser = await this.sp.web.currentUser();
   const email: string = (currentUser.Email || "").toLowerCase();

   if (!email) {
     return "Employee";
   }

   let roles: Array<{
     UserEmail?: string;
     UserName?: string;
     Role?: string;
       isActive?: boolean;
   }> = [];
   try {
     roles = await this.sp.web.lists
       .getByTitle("UserRoles-SAR")
       .items
      .select("UserEmail", "UserName", "Role", "isActive")
      .filter("UserEmail eq '" + email.replace(/'/g, "''") + "' and isActive eq 1")();
   }
   catch (error) {
     console.warn(
       "UserRoles-SAR could not be read. Defaulting the user to Employee.",
       error
     );
     return "Employee";
   }

   const roleRecord: IUserRole | undefined = roles
    .filter((item) => item.isActive === true &&
       (item.Role === "Admin" || item.Role === "HR"))
     .map((item): IUserRole => ({
       UserEmail: item.UserEmail || "",
       UserName: item.UserName || "",
       Role: item.Role === "Admin" ? "Admin" : "HR",
       IsActive: true
    }))[0];

   return roleRecord ? roleRecord.Role : "Employee";
 }

 public async getTrainings(): Promise<ITraining[]> {
   const items = await this.sp.web.lists
     .getByTitle("Trainings-SAR")
     .items
     .select(
       "Id", "Title", "Description", "Category", "Trainer",
       "TrainingDate", "AvailableSeats", "Status"
     )();

   return items.map((item) => ({
     Id: item.Id,
     TrainingName: item.Title || "",
     Description: item.Description || "",
     Category: item.Category || "",
     Trainer: item.Trainer || "",
     TrainingDate: item.TrainingDate || "",
     AvailableSeats: Number(item.AvailableSeats || 0),
     Status: item.Status || ""
   }));
 }

 public async createTraining(training: Omit<ITraining, "Id">, role: UserRole): Promise<void> {
   this.assertTrainingManager(role);
   await this.sp.web.lists.getByTitle("Trainings-SAR").items.add({
     Title: training.TrainingName,
     Description: training.Description,
     Category: training.Category,
     Trainer: training.Trainer,
     TrainingDate: training.TrainingDate,
     AvailableSeats: training.AvailableSeats,
     Status: training.Status
   });
 }

 public async updateTraining(id: number, training: Partial<Omit<ITraining, "Id">>, role: UserRole): Promise<void> {
   this.assertTrainingManager(role);
   await this.sp.web.lists.getByTitle("Trainings-SAR").items.getById(id).update({
     Title: training.TrainingName,
     Description: training.Description,
     Category: training.Category,
     Trainer: training.Trainer,
     TrainingDate: training.TrainingDate,
     AvailableSeats: training.AvailableSeats,
     Status: training.Status
   });
 }

 public async deleteTraining(id: number, role: UserRole): Promise<void> {
   if (!canDeleteTrainings(role)) {
     throw new UnauthorizedError("Only Admin users can delete trainings.");
   }
   await this.sp.web.lists.getByTitle("Trainings-SAR").items.getById(id).delete();
 }

 private assertTrainingManager(role: UserRole): void {
   if (!canManageTrainings(role)) {
     throw new UnauthorizedError("Only Admin and HR users can manage trainings.");
   }
 }
}