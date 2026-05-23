<html>
<head>
	<meta charset="utf-8" />
	<title>Ehazir Shambles - Data Cleaner</title>
</head>
<body>
	<h1>Ehazir Shambles</h1>
	<p>A desktop app for cleaning student data, mapping columns, building ERP-ready records, and exporting class-wise files.</p>

	<section>
		<h2>Folder Tour</h2>
		<table>
			<tr><th>Folder / File</th><th>What It Is</th></tr>
			<tr><td><code>apps/data_cleaner_formatter.py</code></td><td>Main source file for the desktop app.</td></tr>
			<tr><td><code>apps/data_cleaner_exe_app/EhazirShambles.exe</code></td><td>Packaged Windows app you can run without Python.</td></tr>
			<tr><td><code>apps/data_cleaner_exe_app/ehazirshambles.ico</code></td><td>Application icon used by the EXE.</td></tr>
			<tr><td><code>apps/data_cleaner_exe_app/EhazirShambles_logo.png</code></td><td>Splash/logo image shown on startup.</td></tr>
			<tr><td><code>apps/data_cleaner_exe_app/build/</code></td><td>PyInstaller build output used while packaging.</td></tr>
			<tr><td><code>html/</code></td><td>Supporting HTML assets used in the workspace.</td></tr>
			<tr><td><code>.git/</code></td><td>Git repository metadata.</td></tr>
		</table>
	</section>

	<section>
		<h2>How To Download The EXE</h2>
		<ol>
			<li>Open <code>Complete/ehazir/apps/data_cleaner_exe_app/</code>.</li>
			<li>Download or copy <code>EhazirShambles.exe</code> to your Windows machine.</li>
			<li>Keep <code>ehazirshambles.ico</code> and <code>EhazirShambles_logo.png</code> in the same folder if you want the branded assets available during rebuilds.</li>
		</ol>
		<p>If you are viewing this project on GitHub or another file host, download the <code>EhazirShambles.exe</code> file from that folder and run it directly on Windows.</p>
	</section>

	<section>
		<h2>How To Use The App</h2>
		<ol>
			<li>Start the app by double-clicking <code>EhazirShambles.exe</code>.</li>
			<li>Import a source file in XLSX, XLS, or CSV format.</li>
			<li>Review the detected columns and adjust mappings if needed.</li>
			<li>Use the format customizer to set DOB parsing and address handling.</li>
			<li>Clean the data and check the preview.</li>
			<li>Configure school details, ERP options, and presets.</li>
			<li>Separate by class and export the final files.</li>
		</ol>
	</section>

	<section>
		<h2>Features</h2>
		<ul>
			<li>Auto column detection with manual override.</li>
			<li>DOB conversion and missing DOB fallback to <code>0</code>.</li>
			<li>Invalid phone numbers are cleared to plain empty values.</li>
			<li>Class-wise export with roll numbers starting from 1.</li>
			<li>Preset save/load support with rename and delete actions.</li>
			<li>Startup splash branding with icon and logo support.</li>
		</ul>
	</section>

	<section>
		<h2>Rebuild Requirements</h2>
		<p>If you want to rebuild the EXE yourself, use the Python environment in <code>Env1</code> and install these packages:</p>
		<ul>
			<li><code>customtkinter</code></li>
			<li><code>pandas</code></li>
			<li><code>openpyxl</code></li>
			<li><code>pyinstaller</code> if you want to package the app again</li>
		</ul>
	</section>

	<section>
		<h2>Troubleshooting</h2>
		<ul>
			<li>If the app does not maximize correctly, close it and open it again; startup maximize is delayed intentionally.</li>
			<li>If the splash image does not appear, confirm that <code>EhazirShambles_logo.png</code> is present next to the packaged EXE.</li>
			<li>If you rebuild the app, make sure the icon and logo are bundled with PyInstaller.</li>
		</ul>
	</section>
</body>
</html>
