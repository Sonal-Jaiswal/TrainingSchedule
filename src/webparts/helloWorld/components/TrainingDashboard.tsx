// React is used to create the dashboard component and manage its state.
import * as React from "react";

// Fluent UI provides the buttons, fields, messages, and loading indicator.
import {
 PrimaryButton,
 DefaultButton,
 TextField,
 Dropdown,
 IDropdownOption,
 MessageBar,
 MessageBarType,
 Spinner,
 SpinnerSize
} from "@fluentui/react";

// Defines the properties received from the SharePoint web part.
import { IHelloWorldProps } from "./IHelloWorldProps";

// Displays the current user's enrolled courses and progress.
import MyCourses from "./MyCourses";

// Provides the TypeScript models used for SharePoint data.
import {
 IEnrollment,
 ITraining
} from "./TrainingModels";

// Creates a PnPjs client connected to the current SharePoint site.
import { spfi, SPFI, SPFx } from "@pnp/sp";

// Registers the SharePoint APIs used by this component.
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";


// Main component for browsing, enrolling in, and managing trainings.
const HelloWorld: React.FC<IHelloWorldProps> = (props) => {
 // Stores the configured PnPjs client for SharePoint requests.
 const [sp, setSp] = React.useState<SPFI | null>(null);

 // Stores all training records loaded from the Trainings-SAR list.
 const [trainings, setTrainings] =
   React.useState<ITraining[]>([]);

 // Stores all enrollment records loaded from the Enrollments-SAR list.
 const [enrollments, setEnrollments] =
   React.useState<IEnrollment[]>([]);

 // Controls the initial loading state shown while SharePoint data is fetched.
 const [loading, setLoading] =
   React.useState<boolean>(true);

 // Prevents duplicate actions while enrollment or cancellation is processing.
 const [submitting, setSubmitting] =
   React.useState<boolean>(false);

 // Stores an error message displayed in the error message bar.
 const [error, setError] =
   React.useState<string>("");

 // Stores a success message displayed after a completed action.
 const [success, setSuccess] =
   React.useState<string>("");

 // Stores the text entered in the training search field.
 const [searchText, setSearchText] =
   React.useState<string>("");

 // Stores the category currently selected in the category dropdown.
 const [selectedCategory, setSelectedCategory] =
   React.useState<string>("All");

 // Stores the training selected for enrollment confirmation.
 const [selectedTraining, setSelectedTraining] =
   React.useState<ITraining | null>(null);

 // Controls whether the enrollment confirmation modal is visible.
 const [showEnrollForm, setShowEnrollForm] =
   React.useState<boolean>(false);

 // Controls whether My Courses or the training catalog is displayed.
 const [showMyCourses, setShowMyCourses] =
   React.useState<boolean>(false);

 // Stores the display name of the current SharePoint user.
 const [employeeName, setEmployeeName] =
   React.useState<string>("");

 // Controls the compact profile panel shown from the avatar.
 const [showProfile, setShowProfile] =
   React.useState<boolean>(false);

 const employeeInitials: string = employeeName
   .split(" ")
   .filter((namePart: string) => namePart.length > 0)
   .map((namePart: string) => namePart.charAt(0).toUpperCase())
   .slice(0, 2)
   .join("") || "U";

 // Defines the visual layout used by the dashboard and its controls.
 const styles: {
   [key: string]: React.CSSProperties
 } = {
   page: {
     fontFamily: '"Segoe UI", Arial, sans-serif',
     backgroundColor: "#f4f8fb",
     minHeight: "100%",
     padding: "24px",
     boxSizing: "border-box"
   },
   header: {
     background:
       "linear-gradient(135deg, #007f86, #12343b)",
     color: "#ffffff",
     padding: "28px 32px",
     borderRadius: "12px",
     marginBottom: "24px",
     display: "flex",
     justifyContent: "space-between",
    alignItems: "flex-start",
     boxShadow: "0 4px 12px rgba(0,0,0,0.12)"
   },
   headerTitle: {
     margin: "0 0 8px 0",
     fontSize: "28px",
     fontWeight: 600
   },
   headerSubtitle: {
     margin: 0,
     fontSize: "15px",
     opacity: 0.9
   },
   userBadge: {
     display: "flex",
     flexDirection: "column",
     alignItems: "center",
     gap: "10px",
     position: "relative"
   },
   avatarButton: {
     width: "52px",
     height: "52px",
     borderRadius: "50%",
     border: "2px solid rgba(255,255,255,0.85)",
     backgroundColor: "#ffffff",
     color: "#007f86",
     fontSize: "17px",
     fontWeight: 700,
     cursor: "pointer",
     display: "flex",
     alignItems: "center",
     justifyContent: "center",
     boxShadow: "0 2px 8px rgba(0,0,0,0.18)"
   },
   profilePanel: {
     position: "absolute",
     top: "64px",
     right: 0,
     width: "210px",
     padding: "14px",
     backgroundColor: "#ffffff",
     color: "#12343b",
     borderRadius: "8px",
     boxShadow: "0 8px 20px rgba(0,0,0,0.2)",
     zIndex: 10,
     textAlign: "left"
   },
   profileName: {
     margin: 0,
     fontSize: "15px",
     fontWeight: 600,
     overflow: "hidden",
     textOverflow: "ellipsis",
     whiteSpace: "nowrap"
   },
   profileLabel: {
     margin: "5px 0 0",
     color: "#52666d",
     fontSize: "12px"
   },
   statsContainer: {
     display: "grid",
     gridTemplateColumns:
       "repeat(auto-fit, minmax(180px, 1fr))",
     gap: "16px",
     marginBottom: "28px"
   },
   statCard: {
     backgroundColor: "#ffffff",
     borderRadius: "10px",
     padding: "20px",
     boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
     border: "1px solid #e5e5e5"
   },
   statNumber: {
     fontSize: "30px",
     fontWeight: 700,
    color: "#007f86",
     marginBottom: "5px"
   },
   statLabel: {
     color: "#52666d",
     fontSize: "14px"
   },
   section: {
     backgroundColor: "#ffffff",
     borderRadius: "12px",
     padding: "24px",
     marginBottom: "24px",
     boxShadow: "0 2px 8px rgba(0,0,0,0.07)"
   },
   sectionTitle: {
     margin: "0 0 20px 0",
     fontSize: "21px",
    color: "#12343b",
     fontWeight: 600
   },
   filterContainer: {
     display: "grid",
     gridTemplateColumns: "2fr 1fr",
     gap: "20px",
     marginBottom: "24px"
   },
   trainingGrid: {
     display: "grid",
     gridTemplateColumns:
       "repeat(auto-fit, minmax(300px, 1fr))",
     gap: "20px"
   },
   trainingCard: {
     backgroundColor: "#ffffff",
    border: "1px solid #d9e2e7",
     borderRadius: "12px",
     padding: "20px",
     boxShadow: "0 2px 8px rgba(0,0,0,0.06)"
   },
   trainingHeader: {
     display: "flex",
     justifyContent: "space-between",
     alignItems: "flex-start",
     gap: "10px",
     marginBottom: "12px"
   },
   trainingTitle: {
     margin: 0,
     color: "#12343b",
     fontSize: "19px",
     fontWeight: 600
   },
   activeBadge: {
     backgroundColor: "#dff6dd",
     color: "#107c10",
     padding: "5px 10px",
     borderRadius: "15px",
     fontSize: "12px",
     fontWeight: 600,
     whiteSpace: "nowrap"
   },
   inactiveBadge: {
     backgroundColor: "#fde7e9",
     color: "#a4262c",
     padding: "5px 10px",
     borderRadius: "15px",
     fontSize: "12px",
     fontWeight: 600,
     whiteSpace: "nowrap"
   },
   description: {
     color: "#52666d",
     fontSize: "14px",
     lineHeight: "1.5",
     minHeight: "63px",
     marginBottom: "16px"
   },
   details: {
     borderTop: "1px solid #d9e2e7",
     paddingTop: "14px",
     marginBottom: "18px"
   },
   detailRow: {
     display: "flex",
     justifyContent: "space-between",
     gap: "10px",
     padding: "6px 0",
     fontSize: "13px"
   },
   detailLabel: {
     color: "#52666d",
     fontWeight: 600
   },
   detailValue: {
     color: "#12343b",
     textAlign: "right"
   },
   noResults: {
     textAlign: "center",
     padding: "50px 20px",
    color: "#52666d"
   },
   overlay: {
     position: "fixed",
     top: 0,
     left: 0,
     right: 0,
     bottom: 0,
     backgroundColor: "rgba(0,0,0,0.45)",
     display: "flex",
     justifyContent: "center",
     alignItems: "center",
     zIndex: 1000
   },
   modal: {
     backgroundColor: "#ffffff",
     width: "90%",
     maxWidth: "500px",
     borderRadius: "12px",
     padding: "28px",
     boxShadow: "0 10px 30px rgba(0,0,0,0.25)"
   },
   modalTitle: {
     margin: "0 0 20px 0",
     fontSize: "23px",
     color: "#323130"
   },
   modalTraining: {
     backgroundColor: "#edf7f7",
     borderLeft: "4px solid #007f86",
     padding: "15px",
     marginBottom: "20px"
   },
   modalDetail: {
     margin: "8px 0",
     fontSize: "14px",
    color: "#12343b"
   },
   modalButtons: {
     display: "flex",
     gap: "12px",
     marginTop: "25px"
   },
   footer: {
     textAlign: "center",
     color: "#8a8886",
     fontSize: "12px",
     marginTop: "30px"
   }
 };

 // Creates a PnPjs client whenever the SharePoint context changes.
 React.useEffect(() => {
   const spInstance: SPFI =
     spfi().using(
       SPFx(props.context)
     );
   setSp(spInstance);
 }, [props.context]);

 // Loads and maps training records from the Trainings-SAR list.
 const loadTrainings = async (
   spInstance: SPFI
 ): Promise<void> => {
   const items = await spInstance.web.lists
     .getByTitle("Trainings-SAR")
     .items
     .select(
       "Id",
       "Title",
       "Description",
       "Category",
       "Trainer",
       "TrainingDate",
       "AvailableSeats",
       "Status"
     )
     .orderBy(
       "TrainingDate",
       true
     )();
   const data: ITraining[] =
     items.map((item) => ({
       Id: item.Id,
       TrainingName:
         item.Title || "",
       Description:
         item.Description || "",
       Category:
         item.Category || "",
       Trainer:
         item.Trainer || "",
       TrainingDate:
         item.TrainingDate || "",
       AvailableSeats:
         Number(item.AvailableSeats || 0),
       Status:
         item.Status || ""
     }));
   setTrainings(data);
 };

 // Loads enrollment records and expands their employee and training lookups.
 const loadEnrollments = async (
   spInstance: SPFI
 ): Promise<void> => {
   const items = await spInstance.web.lists
     .getByTitle("Enrollments-SAR")
     .items
     .select(
       "Id",
       "Employee/Title",
       "Training/Id",
       "Training/Title",
       "EnrollmentDate",
       "Status",
       "CompletionStatus"
     )
     .expand(
       "Employee",
       "Training"
     )
     .orderBy(
       "Created",
       false
     )();
   const data: IEnrollment[] =
     items.map((item) => ({
       Id: item.Id,
       Employee:
         item.Employee &&
         item.Employee.Title
           ? item.Employee.Title
           : "",
       Training:
         item.Training &&
         item.Training.Title
           ? item.Training.Title
           : "",
       TrainingId:
         item.Training && item.Training.Id
           ? Number(item.Training.Id)
           : 0,
       EnrollmentDate:
         item.EnrollmentDate || "",
       Status:
         item.Status || "",
       CompletionStatus:
         item.CompletionStatus || ""
     }));
   setEnrollments(data);
 };

 // Loads the current user, trainings, and enrollments when SharePoint is ready.
 React.useEffect(() => {
   if (!sp) {
     return;
   }
   const loadData = async (): Promise<void> => {
     try {
       setLoading(true);
       setError("");
       // Current SharePoint user
       const currentUser =
         await sp.web.currentUser();
       setEmployeeName(
         currentUser.Title || ""
       );
       // Load Trainings
       await loadTrainings(sp);
       // Load Enrollments
       await loadEnrollments(sp);
     }
     catch (err) {
       console.error(
         "SharePoint loading error:",
         err
       );
       setError(
         "Unable to load data from SharePoint. Please check the list names and column internal names."
       );
     }
     finally {
       setLoading(false);
     }
   };
    loadData().catch((error) => {
      console.error("Error occurred while loading data:", error);
    });
 }, [sp]);

 // Builds the category dropdown from the categories present in the data.
 const categories: IDropdownOption[] = [
   {
     key: "All",
     text: "All Categories"
   }
 ];
 trainings.forEach(
   (training: ITraining) => {
     if (
       training.Category &&
       categories.filter(
         (category) =>
           category.key ===
           training.Category
       ).length === 0
     ) {
       categories.push({
         key: training.Category,
         text: training.Category
       });
     }
   }
 );

 // Filters trainings by the search text and selected category.
 const filteredTrainings =
   trainings.filter(
     (training: ITraining) => {
       const search: string =
         searchText
           .toLowerCase()
           .trim();
       const trainingName: string =
         training.TrainingName
           .toLowerCase();
       const description: string =
         training.Description
           .toLowerCase();
       const matchesSearch: boolean =
         search === "" ||
         trainingName.indexOf(search) !== -1 ||
         description.indexOf(search) !== -1;
       const matchesCategory: boolean =
         selectedCategory === "All" ||
         training.Category ===
         selectedCategory;
       return (
         matchesSearch &&
         matchesCategory
       );
     }
   );

 // Calculates the summary values shown above the catalog.
 const totalTrainings: number =
   trainings.length;
 const activeTrainings: number =
   trainings.filter(
     (training: ITraining) =>
       training.Status
         .toLowerCase() === "active"
   ).length;
 const totalSeats: number =
   trainings.reduce(
     (
       total: number,
       training: ITraining
     ) =>
       total +
       training.AvailableSeats,
     0
   );
 const myCourses: IEnrollment[] =
   enrollments.filter(
     (enrollment: IEnrollment) =>
       enrollment.Employee === employeeName &&
       enrollment.Status.toLowerCase() !==
       "cancelled"
   );

 // Opens enrollment confirmation unless the training has no available seats.
 const openEnrollment =
   (training: ITraining): void => {
     if (
       training.AvailableSeats <= 0
     ) {
       setError(
         "There are no available seats for this training."
       );
       return;
     }
     setSelectedTraining(training);
     setShowEnrollForm(true);
     setError("");
     setSuccess("");
   };

 // Creates an enrollment and decreases the selected training's seat count.
 const submitEnrollment =
   async (): Promise<void> => {
     if (!sp) {
       return;
     }
     if (!selectedTraining) {
       setError(
         "Please select a training."
       );
       return;
     }
     try {
       setSubmitting(true);
       setError("");
       setSuccess("");
       // ----------------------------------------------------
       // GET CURRENT USER
       // ----------------------------------------------------
       const currentUser =
         await sp.web.currentUser();
       // ----------------------------------------------------
       // CHECK DUPLICATE ENROLLMENT
       // ----------------------------------------------------
       const existing =
         await sp.web.lists
           .getByTitle(
             "Enrollments-SAR"
           )
           .items
           .select(
             "Id",
             "Employee/Id",
             "Training/Id"
           )
           .expand(
             "Employee",
             "Training"
           )
           .filter(
             "Employee/Id eq " +
             currentUser.Id +
             " and Training/Id eq " +
             selectedTraining.Id +
             " and Status ne 'Cancelled'"
           )();
       if (
         existing.length > 0
       ) {
         setError(
           "You are already enrolled in this training."
         );
         return;
       }

       // ----------------------------------------------------
       // GET LATEST TRAINING RECORD
       // ----------------------------------------------------
       const trainingItem =
         await sp.web.lists
           .getByTitle(
             "Trainings-SAR"
           )
           .items
           .getById(
             selectedTraining.Id
           )
           .select(
             "Id",
             "Title",
             "AvailableSeats"
           )();

       const currentSeats: number =
         Number(
           trainingItem.AvailableSeats || 0
         );

       // ----------------------------------------------------
       // CHECK SEATS
       // ----------------------------------------------------
       if (
         currentSeats <= 0
       ) {
         setError(
           "No seats are available for this training."
         );
         return;
       }

       // ----------------------------------------------------
       // CREATE ENROLLMENT
       // ----------------------------------------------------
       await sp.web.lists
         .getByTitle(
           "Enrollments-SAR"
         )
         .items
         .add({
           Title:
             "Enrollment - " +
             currentUser.Title,
           EmployeeId:
             currentUser.Id,
           TrainingId:
             selectedTraining.Id,
           EnrollmentDate:
             new Date().toISOString(),
           Status:
             "Enrolled",
           CompletionStatus:
             "Not Started"
         });

       // ----------------------------------------------------
       // DECREASE AVAILABLE SEATS
       // ----------------------------------------------------
       await sp.web.lists
         .getByTitle(
           "Trainings-SAR"
         )
         .items
         .getById(
           selectedTraining.Id
         )
         .update({
           AvailableSeats:
             currentSeats - 1
         });

       // ----------------------------------------------------
       // RELOAD TRAININGS
       // ----------------------------------------------------
       await loadTrainings(sp);

       // ----------------------------------------------------
       // RELOAD ENROLLMENTS
       // ----------------------------------------------------
       await loadEnrollments(sp);

       // ----------------------------------------------------
       // SUCCESS
       // ----------------------------------------------------
       setSuccess(
         "Successfully enrolled in " +
         selectedTraining.TrainingName +
         "."
       );
       setShowEnrollForm(false);
       setSelectedTraining(null);
     }
     catch (err) {
       console.error(
         "Enrollment error:",
         err
       );
       setError(
         "Enrollment failed. Please verify the SharePoint lookup columns, internal names and permissions."
       );
     }
     finally {
       setSubmitting(false);
     }
   };

 // Cancels an enrollment and returns its seat to the training.
 const cancelCourse =
   async (enrollment: IEnrollment): Promise<void> => {
     if (!sp || enrollment.TrainingId <= 0) {
       return;
     }
     try {
       setSubmitting(true);
       setError("");
       setSuccess("");

       const trainingItem =
         await sp.web.lists
           .getByTitle("Trainings-SAR")
           .items
           .getById(enrollment.TrainingId)
           .select(
             "Id",
             "AvailableSeats",
             "Status"
           )();
       const currentSeats: number =
         Number(trainingItem.AvailableSeats || 0);
       await sp.web.lists
         .getByTitle("Enrollments-SAR")
         .items
         .getById(enrollment.Id)
         .update({
           Status: "Cancelled"
         });

       await sp.web.lists
         .getByTitle("Trainings-SAR")
         .items
         .getById(enrollment.TrainingId)
         .update({
           AvailableSeats: currentSeats + 1
         });

       await loadTrainings(sp);
       await loadEnrollments(sp);
       setSuccess(
         "Your enrollment for " +
         enrollment.Training +
         " has been cancelled."
       );
     }
     catch (err) {
       console.error("Cancellation error:", err);
       setError(
         "Cancellation failed. Please verify your SharePoint permissions."
       );
     }
     finally {
       setSubmitting(false);
     }
   };

 // Closes the enrollment confirmation modal and clears its selection.
 const closeEnrollmentModal =
   (): void => {
     setShowEnrollForm(false);
     setSelectedTraining(null);
     setError("");
   };

 // Show a loading panel until the initial SharePoint request completes.
 if (loading) {
   return (
<div style={styles.page}>
<div
         style={{
           backgroundColor: "#ffffff",
           padding: "50px",
           borderRadius: "12px",
           textAlign: "center"
         }}
>
<Spinner
           size={SpinnerSize.large}
           label={
             "Loading training data from SharePoint..."
           }
         />
</div>
</div>
   );
 }

 // Render the dashboard after the initial data load is complete.
 return (
<div style={styles.page}>
     {/* ================================================== */}
     {/* HEADER */}
     {/* ================================================== */}
<div style={styles.header}>
<div>
<h1 style={styles.headerTitle}>
           Employee Training Dashboard
</h1>
<p style={styles.headerSubtitle}>
           Discover and enroll in internal
           training programs
</p>
</div>
<div style={styles.userBadge}>
<button
         type="button"
         aria-label="Open profile"
         title={employeeName || "Open profile"}
         onClick={() => setShowProfile(!showProfile)}
         style={styles.avatarButton}
       >
         {employeeInitials}
</button>
       {showProfile && (
<div style={styles.profilePanel}>
<p style={styles.profileName}>
           {employeeName || "User"}
</p>
<p style={styles.profileLabel}>
           SharePoint employee
</p>
</div>
       )}
<DefaultButton
         text={showMyCourses ? "All Trainings" : "My Courses"}
         onClick={() => setShowMyCourses(!showMyCourses)}
         styles={{
           root: {
             borderRadius: "6px"
           },
           rootHovered: {
             borderRadius: "6px"
           }
         }}
       />
</div>
    </div>

     {/* ================================================== */}
     {/* ERROR */}
     {/* ================================================== */}
     {error && (
<div
         style={{
           marginBottom: "20px"
         }}
>
<MessageBar
           messageBarType={
             MessageBarType.error
           }
           isMultiline={true}
           onDismiss={() =>
             setError("")
           }
>
           {error}
</MessageBar>
</div>
     )}

     {/* ================================================== */}
     {/* SUCCESS */}
     {/* ================================================== */}
     {success && (
<div
         style={{
           marginBottom: "20px"
         }}
>
<MessageBar
           messageBarType={
             MessageBarType.success
           }
           onDismiss={() =>
             setSuccess("")
           }
>
           {success}
</MessageBar>
</div>
     )}

     {/* ================================================== */}
     {/* STATISTICS */}
     {/* ================================================== */}
<div style={styles.statsContainer}>
<div style={styles.statCard}>
<div style={styles.statNumber}>
           {totalTrainings}
</div>
<div style={styles.statLabel}>
           Total Available Trainings
</div>
</div>

<div style={styles.statCard}>
<div style={styles.statNumber}>
           {activeTrainings}
</div>
<div style={styles.statLabel}>
           Active Trainings
</div>
</div>

<div style={styles.statCard}>
<div style={styles.statNumber}>
           {totalSeats}
</div>
<div style={styles.statLabel}>
           Available Seats
</div>
</div>

</div>

     {/* ================================================== */}
     {/* MY COURSES */}
     {/* ================================================== */}
{showMyCourses && <div style={styles.section}>
       <MyCourses
         courses={myCourses}
         submitting={submitting}
         onCancel={cancelCourse}
       />
</div>}

     {/* ================================================== */}
     {/* TRAININGS */}
     {/* ================================================== */}
{!showMyCourses && <div style={styles.section}>
<h2 style={styles.sectionTitle}>
         Available Trainings
</h2>

       {/* SEARCH + CATEGORY */}
<div style={styles.filterContainer}>
<TextField
           label="Search Training"
           placeholder={
             "Search by training name or description"
           }
           value={searchText}
           onChange={
             (
               _event,
               value
             ) =>
               setSearchText(
                 value || ""
               )
           }
         />

<Dropdown
           label="Category"
           selectedKey={
             selectedCategory
           }
           options={categories}
           onChange={
             (
               _event,
               option
             ) => {
               if (option) {
                 setSelectedCategory(
                   option.key as string
                 );
               }
             }
           }
         />
</div>

       {/* TRAINING CARDS */}
<div style={styles.trainingGrid}>
         {filteredTrainings.map(
           (
             training: ITraining
           ) => (
<div
               key={training.Id}
               style={styles.trainingCard}
>
               {/* CARD HEADER */}
<div
                 style={styles.trainingHeader}
>
<h3
                   style={
                     styles.trainingTitle
                   }
>
                   {training.TrainingName}
</h3>

<span
                   style={
                     training.Status
                       .toLowerCase() ===
                     "active"
                       ? styles.activeBadge
                       : styles.inactiveBadge
                   }
>
                   {training.Status}
</span>
</div>

               {/* DESCRIPTION */}
<p
                 style={
                   styles.description
                 }
>
                 {training.Description}
</p>

               {/* DETAILS */}
<div style={styles.details}>
<div
                   style={styles.detailRow}
>
<span
                     style={
                       styles.detailLabel
                     }
>
                     Category
</span>
<span
                     style={
                       styles.detailValue
                     }
>
                     {training.Category}
</span>
</div>

<div
                   style={styles.detailRow}
>
<span
                     style={
                       styles.detailLabel
                     }
>
                     Trainer
</span>
<span
                     style={
                       styles.detailValue
                     }
>
                     {training.Trainer}
</span>
</div>

<div
                   style={styles.detailRow}
>
<span
                     style={
                       styles.detailLabel
                     }
>
                     Training Date
</span>
<span
                     style={
                       styles.detailValue
                     }
>
                     {
                       training.TrainingDate
                         ? new Date(
                             training.TrainingDate
                           ).toLocaleDateString()
                         : "Not specified"
                     }
</span>
</div>

<div
                   style={styles.detailRow}
>
<span
                     style={
                       styles.detailLabel
                     }
>
                     Available Seats
</span>
<span
                     style={{
                       ...styles.detailValue,
                       fontWeight: 600
                     }}
>
                     {
                       training.AvailableSeats
                     }
</span>
</div>
</div>

               {/* ENROLL BUTTON */}
<PrimaryButton
                 text={
                   training.AvailableSeats > 0
                     ? "Enroll Now"
                     : "Fully Booked"
                 }
                 disabled={
                   training.AvailableSeats <= 0
                 }
                 onClick={() =>
                   openEnrollment(
                     training
                   )
                 }
                 styles={{
                   root: {
                     width: "100%",
                     borderRadius: "6px"
                   }
                 }}
               />
</div>
           )
         )}
</div>

       {/* NO RESULTS */}
       {filteredTrainings.length === 0 && (
<div
           style={styles.noResults}
>
<h3>
             No trainings found
</h3>
<p>
             Try changing your search
             or category.
</p>
</div>
       )}
    </div>}

     {/* ================================================== */}
     {/* ENROLLMENT MODAL */}
     {/* ================================================== */}
     {showEnrollForm &&
       selectedTraining && (
<div
           style={styles.overlay}
>
<div
             style={styles.modal}
>
<h2
               style={
                 styles.modalTitle
               }
>
               Confirm Enrollment
</h2>

             {/* TRAINING INFORMATION */}
<div
               style={
                 styles.modalTraining
               }
>
<h3
                 style={{
                   margin:
                     "0 0 12px 0",
                   color: "#007f86"
                 }}
>
                 {
                   selectedTraining.TrainingName
                 }
</h3>

<p
                 style={
                   styles.modalDetail
                 }
>
<strong>
                   Category:
</strong>{" "}
                 {
                   selectedTraining.Category
                 }
</p>

<p
                 style={
                   styles.modalDetail
                 }
>
<strong>
                   Trainer:
</strong>{" "}
                 {
                   selectedTraining.Trainer
                 }
</p>

<p
                 style={
                   styles.modalDetail
                 }
>
<strong>
                   Date:
</strong>{" "}
                 {
                   selectedTraining.TrainingDate
                     ? new Date(
                         selectedTraining.TrainingDate
                       ).toLocaleDateString()
                     : "Not specified"
                 }
</p>

<p
                 style={
                   styles.modalDetail
                 }
>
<strong>
                   Available Seats:
</strong>{" "}
                 {
                   selectedTraining.AvailableSeats
                 }
</p>
</div>

             {/* EMPLOYEE */}
<p
               style={
                 styles.modalDetail
               }
>
<strong>
                 Employee:
</strong>{" "}
               {employeeName}
</p>

<p
               style={{
                 fontSize: "13px",
                 color: "#52666d"
               }}
>
               Your enrollment will be saved
               directly to the SharePoint
               Enrollments-SAR list.
</p>

             {/* BUTTONS */}
<div
               style={
                 styles.modalButtons
               }
>
<PrimaryButton
                 text={
                   submitting
                     ? "Submitting..."
                     : "Confirm Enrollment"
                 }
                 disabled={submitting}
                 onClick={
                   submitEnrollment
                 }
               />

<DefaultButton
                 text="Cancel"
                 disabled={submitting}
                 onClick={
                   closeEnrollmentModal
                 }
               />
</div>
</div>
</div>
       )}

     {/* ================================================== */}
     {/* FOOTER */}
     {/* ================================================== */}
<div style={styles.footer}>
       Training data powered by SharePoint & PnPjs
</div>
</div>
 );
};

// ============================================================
// EXPORT
// ============================================================
export default HelloWorld;