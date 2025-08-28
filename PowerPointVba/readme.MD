# Orchestration
## Why I Tried and Then Abandoned PowerShell for Orchestrating VBA Generation

At first, I thought PowerShell would be a great choice for managing the batch-generation of VBA objects, especially for creating multiple PowerPoint presentations from a template. Since both PowerShell and VBA are Microsoft technologies, I figured they would work together without any issues. Turns out, I was wrong.

The biggest headache was debugging. While I could debug PowerShell scripts and VBA macros on their own, trying to troubleshoot the workflow between the two was a nightmare. It was frustrating and took way too much time.

On top of that, VBA is much better integrated with PowerPoint. It lets you directly work with slides, shapes, and presentations without needing to rely on external scripts. Keeping everything inside PowerPoint made the workflow simpler, got rid of extra dependencies, and made it easier to share the solution. By embedding the VBA code directly into the PowerPoint file, I didn’t have to worry about users installing or setting up PowerShell.

In the end, PowerShell is a powerful tool, but it just wasn’t the right fit for this job. VBA turned out to be a cleaner, more integrated, and easier-to-maintain solution for automating PowerPoint generation.