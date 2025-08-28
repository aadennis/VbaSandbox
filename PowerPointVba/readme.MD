# Orchestration
## Why I Tried and Then Abandoned PowerShell for Orchestrating VBA Generation

At first, I thought PowerShell would be a good choice for orchestrating the batch-generation of VBA objects, particularly for creating multiple PowerPoint presentations using a template. After all, both PowerShell and VBA are flagship Microsoft technologies, so I assumed they would work seamlessly together. Experience proved me wrong.

The biggest challenge was debugging. While I could debug PowerShell scripts and VBA macros individually, debugging the integrated workflow between the two was impractical. This made troubleshooting errors frustrating and time-consuming.

Plus, VBA has tighter integration with the PowerPoint object model, allowing direct manipulation of slides, shapes, and presentations. By keeping all the logic within the PowerPoint environment, I could simplify the workflow, avoid external dependencies, and make the solution easier to distribute. Embedding the VBA code directly into the PowerPoint file also meant that users don't need to install or configure PowerShell.

In the end, while PowerShell is a powerful tool, it wasn’t the right fit for this task. VBA provided a more streamlined, integrated, and maintainable solution for automating PowerPoint generation.
