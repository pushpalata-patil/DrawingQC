DrawingQC Local Agent
=====================

What it is
----------
A small helper that runs on YOUR Windows PC and lets the Support Automation website
(hosted on Render) use the programs installed on your machine:

  * AutoCAD / AutoCAD Plant 3D  -> "Sync to AutoCAD" writes the QC register into the open drawing
  * Microsoft Word              -> "Generate booklet" fills the template and exports the PDF

The website never sends your files anywhere else: the booklet is built on your PC and
downloaded from the agent on http://127.0.0.1:5081.

How to use
----------
1. Unzip this folder anywhere (e.g. C:\DrawingQC-Agent).
2. Double-click DrawingQC.Agent.exe. A console window stays open while it runs:
       DrawingQC Local Agent 1.0.0 listening on http://127.0.0.1:5081
   (Windows SmartScreen may ask once: "More info" -> "Run anyway".)
3. Open the website. The sidebar shows "Local agent: connected" with the AutoCAD
   product it found. Use the tools as normal.
4. Close the console window to stop the agent.

Notes
-----
* Requires nothing else: .NET is bundled inside the .exe.
* Keep AutoCAD open with a drawing before clicking "Sync to AutoCAD".
* Microsoft Word must be installed for the booklet.
* Chrome/Edge may show a one-time "allow this site to access your local network" prompt
  the first time the website contacts the agent. Click Allow.

Options (command line)
----------------------
  DrawingQC.Agent.exe --port 5081
  DrawingQC.Agent.exe --origin https://your-site.example.com   (allow another site; repeatable)

By default the agent accepts requests from localhost and any *.onrender.com site.
Environment variables DRAWINGQC_AGENT_PORT and DRAWINGQC_AGENT_ORIGINS (comma-separated)
do the same thing.
