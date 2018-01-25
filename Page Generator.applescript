tell application "Finder"
	set all_servers to (every disk whose name starts with "Server")
	set server to the first item of all_servers
	set server_path to the POSIX path of (server as alias)
end tell

set python_preamble to "LC_ALL=en_GB.utf-8 /usr/local/bin/python3 "
set masters_dir to server_path & "Production\\ Resources/Master\\ pages/"
set python_script to masters_dir & "ms-py-indesign/gen.py"
set master_file to masters_dir & "2018\\ Master.indd"
set pages_dir to masters_dir & "Fresh\\ pages/"

tell application "Adobe InDesign CS4.app"
	activate
end tell

do shell script python_preamble & python_script & " --master=" & master_file & " --pages_dir=" & pages_dir

tell application "Finder"
	reveal POSIX file (server_path & "Production Resources/Master pages/Fresh pages/") as alias
	activate
end tell
