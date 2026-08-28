import { getSP } from './pnpjsConfig';
export interface ITraining {
 Id: number;
 Title: string;
 Description: string;
 Category: string;
 Trainer: string;
 TrainingDate: string;
 AvailableSeats: number;
 Status: string;
}
export async function getTrainings(this: any): Promise<ITraining[]> {
 const sp = getSP(this.context);
 const trainings = await sp.web.lists
   .getByTitle('Trainings')
   .items
   .select(
     'Id',
     'Title',
     'Description',
     'Category',
     'Trainer',
     'TrainingDate',
     'AvailableSeats',
     'Status'
   )();
 return trainings as ITraining[];
}