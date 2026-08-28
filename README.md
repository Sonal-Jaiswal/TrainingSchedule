# SPFx

## Summary

Short summary on functionality and used technologies.

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

## System Architecture

This solution is a SharePoint Framework client-side web part built with React,
Fluent UI, TypeScript, and PnPjs.

```text
Employee
   |
SharePoint Page
   |
HelloWorldWebPart
   |
WelcomePage -> TrainingDashboard -> MyCourses
                     |
                 PnPjs / SPFx context
                  /              \
        Trainings-SAR        Enrollments-SAR
```

### Application Layers

- **Web part layer:** `HelloWorldWebPart.ts` is the SPFx entry point. It mounts
  React and switches between the welcome page and training dashboard.
- **Presentation layer:** `WelcomePage.tsx`, `TrainingDashboard.tsx`, and
  `MyCourses.tsx` provide the user interface and interactions.
- **Data access layer:** `pnpjsConfig.ts` creates a PnPjs client using the
  current SharePoint `WebPartContext`. The dashboard reads and updates list
  data through PnPjs.
- **Data layer:** SharePoint lists store trainings and employee enrollments.
- **Demo progress layer:** ten course modules are generated for each
  enrollment. Module completion is saved in browser `localStorage` using the
  `training-course-module-progress` key.

### SharePoint Lists

#### Trainings-SAR

`Id`, `Title`, `Description`, `Category`, `Trainer`, `TrainingDate`,
`AvailableSeats`, and `Status`.

#### Enrollments-SAR

`Id`, `Employee`, `Training`, `EnrollmentDate`, `Status`, and
`CompletionStatus`.

### Main Workflows

**Loading:** The current user, trainings, and enrollments are loaded when the
dashboard initializes. The dashboard filters enrollments for the current user.

**Enrollment:** The application checks for an existing non-cancelled
enrollment, verifies the latest seat count, creates an enrollment record, and
decreases `AvailableSeats`.

**Cancellation:** The enrollment status is changed to `Cancelled`, one seat is
returned to the related training, and the dashboard data is refreshed.

**Module progress:** Users can check or uncheck ten demo modules per course.
The graph, progress bars, and summary cards update immediately. Progress stays
after a browser refresh until the user changes the module selections.

> Module progress is currently browser-local and is not shared between devices.
> For enterprise persistence, create a SharePoint module-progress list.

## Project Structure

```text
demo/
├── config/                  # SPFx build, serve, package, and deployment config
├── src/
│   ├── index.ts
│   └── webparts/helloWorld/
│       ├── HelloWorldWebPart.ts
│       ├── HelloWorldWebPart.manifest.json
│       ├── assets/
│       ├── components/
│       │   ├── IHelloWorldProps.ts
│       │   ├── MyCourses.tsx
│       │   ├── pnpjsConfig.ts
│       │   ├── TrainingDashboard.tsx
│       │   ├── TrainingModels.ts
│       │   ├── TrainingService.ts
│       │   ├── TrainingCard.tsx
│       │   └── WelcomePage.tsx
│       └── loc/              # Localized strings
├── lib/                      # Compiled JavaScript and declarations
├── release/                  # Bundles, manifests, and audit output
├── teams/                    # Teams packaging output
├── temp/                     # Temporary SPFx build output
├── gulpfile.js               # Gulp/SPFx task configuration
├── package.json              # Dependencies and npm scripts
├── tsconfig.json             # TypeScript compiler configuration
└── README.md
```

## Development and Deployment

```bash
npm install
npm run build
gulp serve
```

- SPFx version: `1.21.0`
- React version: `17.0.1`
- Fluent UI version: `8.106.4`
- PnPjs version: `4.21.0`
- TypeScript target: `ES5`
- Package output: `solution/rishi.sppkg`
- Development workbench: `config/serve.json`

The dashboard uses the `Trainings-SAR` list. The older `TrainingService.ts`
helper currently references `Trainings`; align that list name if the helper is
used again.

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
