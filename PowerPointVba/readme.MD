# Orchestration
## Why I tried and then abandoned PowerShell for the orchestration of VBA generation

I had naively thought that PowerShell would be useful for orchestrating the batch-generation of VBA objects, specifically multiple PowerPoint objects using a template. My thinking was that, well, they are both flagship Microsoft products. However, they simply do not play nicely together, and it is not possible, in any practical way, to debug from PowerShell into VBA. So I gave up.

Putting some specifics on that:
Tighter Integration with PowerPoint: VBA provides direct integration with PowerPoint's object model, making it easier to manipulate presentations, slides, and shapes without needing to rely on external scripts or COM interop.

Simplified Workflow: Using VBA allows you to keep all the logic and automation within the PowerPoint environment, reducing the need for external dependencies or additional setup steps.

Ease of Distribution: A VBA-based solution can be embedded directly into a PowerPoint file, making it easier to share and execute without requiring users to install or configure PowerShell.

Error Handling and Debugging: Debugging VBA code within the PowerPoint environment is more straightforward than troubleshooting PowerShell scripts: I can debug PowerShell, I can debug VBA. But debugging the integrated whole? Nah.

