import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import { IEnrollment } from "./TrainingModels";

// A course enrollment displayed in the user's course list.
export type IMyCourse = IEnrollment;

// Values supplied by the parent dashboard component.
export interface IMyCoursesProps {
 courses: IMyCourse[];
 submitting: boolean;
 onCancel: (course: IMyCourse) => Promise<void>;
}

// Converts the SharePoint completion status into a progress percentage.
const getProgressPercentage = (completionStatus: string): number => {
 const normalizedStatus: string = completionStatus.toLowerCase().trim();

 if (normalizedStatus === "completed" || normalizedStatus === "complete") {
   return 100;
 }

 if (
   normalizedStatus === "in progress" ||
   normalizedStatus === "in-progress" ||
   normalizedStatus === "started"
 ) {
   return 50;
 }

 return 0;
};

const MODULE_COUNT: number = 10;
const MODULE_INDEXES: number[] = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9];
const MODULE_STORAGE_KEY: string = "training-course-module-progress";

const createInitialModules = (completionStatus: string): boolean[] => {
 const progress: number = getProgressPercentage(completionStatus);
 const completedModules: number = Math.round(
   (progress / 100) * MODULE_COUNT
 );
 const modules: boolean[] = [];

 for (let index: number = 0; index < MODULE_COUNT; index++) {
   modules.push(index < completedModules);
 }

 return modules;
};

// Displays the courses enrolled by the current SharePoint user.
const MyCourses: React.FC<IMyCoursesProps> = (props) => {
 const [selectedCourseId, setSelectedCourseId] =
   React.useState<number | null>(null);
 const [completedModules, setCompletedModules] =
   React.useState<{ [courseId: number]: boolean[] }>({});

 React.useEffect(() => {
   setCompletedModules((currentModules) => {
     const nextModules: { [courseId: number]: boolean[] } = {
       ...currentModules
     };

     let savedModules: { [courseId: number]: boolean[] } = {};
     try {
       const savedValue: string | null =
         window.localStorage.getItem(MODULE_STORAGE_KEY);
       if (savedValue) {
         savedModules = JSON.parse(savedValue) as {
           [courseId: number]: boolean[]
         };
       }
     }
     catch (error) {
       console.error("Unable to restore module progress:", error);
     }

     props.courses.forEach((course: IMyCourse) => {
       if (!nextModules[course.Id]) {
         nextModules[course.Id] = savedModules[course.Id] &&
           savedModules[course.Id].length === MODULE_COUNT
           ? savedModules[course.Id]
           : createInitialModules(course.CompletionStatus);
       }
     });

     return nextModules;
   });
 }, [props.courses]);

 const selectedCourse: IMyCourse | undefined =
   props.courses.filter(
     (course: IMyCourse) => course.Id === selectedCourseId
   )[0];

 const getCourseProgress = (course: IMyCourse): number => {
   const modules: boolean[] = completedModules[course.Id];
   if (!modules) {
     return getProgressPercentage(course.CompletionStatus);
   }

   return Math.round(
     (modules.filter((isComplete: boolean) => isComplete).length /
       MODULE_COUNT) * 100
   );
 };

 const completedCourses: number = props.courses.filter(
   (course: IMyCourse) =>
     getCourseProgress(course) === 100
 ).length;

 const averageProgress: number = props.courses.length === 0
   ? 0
   : Math.round(
       props.courses.reduce(
         (total: number, course: IMyCourse) =>
           total + getCourseProgress(course),
         0
       ) / props.courses.length
     );

 const toggleModule = (course: IMyCourse, moduleIndex: number): void => {
   setSelectedCourseId(course.Id);
   setCompletedModules((currentModules) => {
     const modules: boolean[] = currentModules[course.Id] ||
       createInitialModules(course.CompletionStatus);
     const updatedModules: boolean[] = modules.slice();
     updatedModules[moduleIndex] = !updatedModules[moduleIndex];
     const nextModules: { [courseId: number]: boolean[] } = {
       ...currentModules,
       [course.Id]: updatedModules
     };
     try {
       window.localStorage.setItem(
         MODULE_STORAGE_KEY,
         JSON.stringify(nextModules)
       );
     }
     catch (error) {
       console.error("Unable to save module progress:", error);
     }
     return {
       ...nextModules
     };
   });
 };

 // Show a confirmation before changing SharePoint data.
 const handleCancel = (course: IMyCourse): void => {
   if (!window.confirm("Cancel this course enrollment?")) {
     return;
   }

   props.onCancel(course).catch((error) => {
     console.error("Error cancelling enrollment:", error);
   });
 };

 return (
<div>
<h2>
       My Courses
</h2>

     {props.courses.length > 0 && (
<div style={{
       display: "grid",
       gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))",
       gap: "12px",
       marginBottom: "24px"
     }}>
<div style={{
         backgroundColor: "#edf7f7",
         borderLeft: "4px solid #007f86",
         padding: "14px 16px"
       }}>
<div style={{ color: "#52666d", fontSize: "12px" }}>
         Enrolled Courses
</div>
<strong style={{ color: "#12343b", fontSize: "24px" }}>
         {props.courses.length}
</strong>
</div>
<div style={{
         backgroundColor: "#f3f8fd",
         borderLeft: "4px solid #0078d4",
         padding: "14px 16px"
       }}>
<div style={{ color: "#52666d", fontSize: "12px" }}>
         Average Progress
</div>
<strong style={{ color: "#12343b", fontSize: "24px" }}>
         {averageProgress}%
</strong>
</div>
<div style={{
         backgroundColor: "#f1faf0",
         borderLeft: "4px solid #107c10",
         padding: "14px 16px"
       }}>
<div style={{ color: "#52666d", fontSize: "12px" }}>
         Completed
</div>
<strong style={{ color: "#12343b", fontSize: "24px" }}>
         {completedCourses}
</strong>
</div>
</div>
     )}

     {props.courses.length > 0 && (
<div style={{
       border: "1px solid #d9e2e7",
       padding: "18px",
       marginBottom: "24px",
       backgroundColor: "#fbfdfd"
     }}>
<div style={{
         display: "flex",
         justifyContent: "space-between",
         alignItems: "baseline",
         gap: "12px",
         marginBottom: "16px"
       }}>
<h3 style={{ margin: 0, color: "#12343b", fontSize: "17px" }}>
         Progress Overview
</h3>
<span style={{ color: "#52666d", fontSize: "12px" }}>
         Select a course for details
</span>
</div>
<div style={{
         display: "flex",
         alignItems: "flex-end",
         gap: "12px",
         height: "150px",
         overflowX: "auto",
         padding: "8px 4px 0"
       }}>
         {props.courses.map((course: IMyCourse) => {
           const progress: number = getCourseProgress(course);
           const isSelected: boolean = course.Id === selectedCourseId;
           return (
<button
             key={course.Id}
             type="button"
             aria-label={`${course.Training}: ${progress}% complete`}
             onClick={() => setSelectedCourseId(course.Id)}
             style={{
               border: isSelected ? "2px solid #007f86" : "1px solid #d9e2e7",
               backgroundColor: isSelected ? "#edf7f7" : "#ffffff",
               minWidth: "74px",
               height: "134px",
               padding: "8px 6px 6px",
               cursor: "pointer",
               display: "flex",
               flexDirection: "column",
               justifyContent: "flex-end",
               alignItems: "center"
             }}
           >
<span style={{
               color: "#12343b",
               fontSize: "12px",
               fontWeight: 600,
               marginBottom: "6px"
             }}>
             {progress}%
</span>
<span style={{
               width: "28px",
               height: `${Math.max(progress, 4)}px`,
               backgroundColor: progress === 100 ? "#107c10" : "#0078d4",
               transition: "height 250ms ease"
             }} />
<span style={{
               color: "#52666d",
               fontSize: "11px",
               marginTop: "7px",
               maxWidth: "64px",
               overflow: "hidden",
               textOverflow: "ellipsis",
               whiteSpace: "nowrap"
             }}>
             {course.Training}
</span>
</button>
           );
         })}
</div>
         {selectedCourse && (
<div style={{
           marginTop: "16px",
           padding: "12px 14px",
           backgroundColor: "#edf7f7",
           color: "#12343b",
           fontSize: "13px"
         }}>
<strong>{selectedCourse.Training}</strong>
           {" - "}
           {getCourseProgress(selectedCourse)}% complete
</div>
         )}
</div>
     )}

     {/* Show an empty state when the user has no active enrollments. */}
     {props.courses.length === 0 ? (
<div style={{
         textAlign: "center",
         padding: "50px 20px",
         color: "#605e5c"
       }}>
       <p>You are not enrolled in any courses.</p>
</div>
     ) : (
<div style={{
       display: "grid",
       gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))",
       gap: "20px"
     }}>
       {props.courses.map((course: IMyCourse) => (
<div
         key={course.Id}
         style={{
           backgroundColor: "#ffffff",
           border: "1px solid #e1dfdd",
           borderRadius: "12px",
           padding: "20px",
           boxShadow: "0 2px 8px rgba(0,0,0,0.06)"
         }}
>
<h3 style={{
           margin: "0 0 12px 0",
           color: "#323130",
           fontSize: "19px",
           fontWeight: 600
         }}>
           {course.Training}
</h3>

       {/* Show the user's current completion progress. */}
       <div style={{ marginBottom: "18px" }}>
<div style={{
           display: "flex",
           justifyContent: "space-between",
           marginBottom: "6px",
           fontSize: "13px"
         }}>
<span style={{ color: "#605e5c", fontWeight: 600 }}>
           Course Progress
</span>
<span style={{ color: "#323130", fontWeight: 600 }}>
           {getCourseProgress(course)}%
</span>
</div>
<div
         role="progressbar"
         aria-label={`${course.Training} progress`}
         aria-valuemin={0}
         aria-valuemax={100}
         aria-valuenow={getCourseProgress(course)}
         style={{
           height: "10px",
           backgroundColor: "#edebe9",
           borderRadius: "5px",
           overflow: "hidden"
         }}
       >
<div style={{
           width: `${getCourseProgress(course)}%`,
           height: "100%",
           backgroundColor: getCourseProgress(course) === 100
             ? "#107c10"
             : "#0078d4",
           transition: "width 300ms ease"
         }} />
</div>
</div>

<div style={{
           borderTop: "1px solid #edebe9",
           paddingTop: "14px",
           marginBottom: "18px"
         }}>
<div style={{
             display: "flex",
             justifyContent: "space-between",
             alignItems: "center",
             marginBottom: "10px"
           }}>
<strong style={{ color: "#12343b", fontSize: "14px" }}>
             Course Modules
</strong>
<span style={{ color: "#52666d", fontSize: "12px" }}>
             {completedModules[course.Id]
               ? completedModules[course.Id].filter((isComplete: boolean) => isComplete).length
               : 0} / {MODULE_COUNT} complete
</span>
</div>
<div style={{
             display: "grid",
             gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
             gap: "7px"
           }}>
           {MODULE_INDEXES.map((moduleIndex: number) => {
             const isComplete: boolean = Boolean(
               completedModules[course.Id] &&
               completedModules[course.Id][moduleIndex]
             );
             return (
<label
               key={moduleIndex}
               style={{
                 display: "flex",
                 alignItems: "center",
                 gap: "7px",
                 padding: "7px 8px",
                 backgroundColor: isComplete ? "#f1faf0" : "#f8fafb",
                 color: isComplete ? "#107c10" : "#52666d",
                 fontSize: "12px",
                 cursor: "pointer"
               }}
             >
<input
                 type="checkbox"
                 checked={isComplete}
                 onChange={() => toggleModule(course, moduleIndex)}
               />
               Module {moduleIndex + 1}
</label>
             );
           })}
</div>
</div>

<div style={{
           borderTop: "1px solid #edebe9",
           paddingTop: "14px",
           marginBottom: "18px"
         }}>
<div style={{
             display: "flex",
             justifyContent: "space-between",
             gap: "10px",
             padding: "6px 0",
             fontSize: "13px"
           }}>
<span style={{ color: "#605e5c", fontWeight: 600 }}>
             Enrolled On
</span>
<span style={{ color: "#323130", textAlign: "right" }}>
             {course.EnrollmentDate
               ? new Date(course.EnrollmentDate).toLocaleDateString()
               : "Not specified"}
</span>
</div>

<div style={{
             display: "flex",
             justifyContent: "space-between",
             gap: "10px",
             padding: "6px 0",
             fontSize: "13px"
           }}>
<span style={{ color: "#605e5c", fontWeight: 600 }}>
             Completion
</span>
<span style={{ color: "#323130", textAlign: "right" }}>
             {course.CompletionStatus}
</span>
</div>
</div>

<DefaultButton
         text="Cancel Course"
         disabled={props.submitting}
         onClick={() => handleCancel(course)}
         styles={{
           root: {
             width: "100%",
             borderRadius: "6px"
           }
         }}
       />
</div>
       ))}
</div>
     )}
</div>
 );
};

export default MyCourses;
