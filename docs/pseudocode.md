## Outlook Attachments Export Tool – Main Control Flow (Pseudocode)

### Pseudocode

1. **Parse Command-Line Arguments**
	- Define and parse arguments for output directory, date range, sender(s), subject/body keywords, attachment presence/absence, folder path, limit, batch size, dry-run, open-folder, log file, verbosity, etc.
	- Store parsed arguments in `args`.

2. **Set Up Logging**
	- Determine log level based on verbosity/quiet flags.
	- Configure logging handlers (console, optional file).
	- Set log format.

3. **Main Execution**
	- Log start message.
	- If both `--with-attachments` and `--without-attachments` are set:
		- Log error and exit with code 2.

	- **Try:**
		1. **Process Messages**
			- Connect to Outlook via COM.
			- Get Outlook namespace.
			- Get target folder (default Inbox or traverse custom path).
			- Ensure output directory exists.
			- Initialize `seen_hashes`, `saved_count`, `matched_messages`.
			- **Iterate Messages in Batches:**
				- For each batch of messages:
					- For each message:
						- **Filter Message:**
							- Check date range, sender, subject/body keywords, attachment presence/absence.
							- If not matching, continue.
						- Increment `matched_messages`.
						- Get attachments collection.
						- If no attachments and `--with-attachments`, continue.
						- If attachments exist and `--without-attachments`, continue.
						- **Iterate Attachments:**
							- For each attachment:
								- Skip inline unless requested.
								- Determine target directory (sanitize, truncate, fallback if needed).
								- Determine filename (sanitize, truncate).
								- If dry-run: log intended save path, continue.
								- Save attachment to temp file (handle errors).
								- Compute hash of temp file.
								- **Duplicate Detection:**
									- If hash seen before: save to duplicates subfolder, log, update `seen_hashes`.
									- Else: save to final path, log, update `seen_hashes`.
								- Increment `saved_count` if saved.
			- Log number of matched messages and saved attachments.
			- Return `saved_count`.

		2. **If `--open-folder` is set:**
			- Try to open output folder in file explorer (handle errors).

		3. Log completion and number of saved attachments.
		4. Exit with code 0.

	- **Except Exception:**
		- Log unhandled error with traceback.
		- Exit with code 1.

4. **END PROGRAM**

---

### Mermaid Flowchart

```mermaid
flowchart TD
	A([Start]) --> B[Parse command-line arguments]
	B --> C[Set up logging]
	C --> D{Both --with-attachments and --without-attachments?}
	D -- Yes --> E[Log error, exit code 2]
	D -- No --> F[Connect to Outlook, get folder]
	F --> G[Ensure output directory exists]
	G --> H[Initialize seen_hashes, counters]
	H --> I[Iterate messages in batches]
	I --> J{Message matches filters?}
	J -- No --> I
	J -- Yes --> K[Get attachments]
	K --> L{Attachment count == 0 and --with-attachments?}
	L -- Yes --> I
	L -- No --> M{Attachment count > 0 and --without-attachments?}
	M -- Yes --> I
	M -- No --> N[Iterate attachments]
	N --> O{Skip inline?}
	O -- Yes --> N
	O -- No --> P[Determine target dir, filename]
	P --> Q{Dry-run?}
	Q -- Yes --> N
	Q -- No --> R[Save attachment to temp file]
	R --> S[Compute hash]
	S --> T{Duplicate?}
	T -- Yes --> U[Save to duplicates subfolder]
	T -- No --> V[Save to final path]
	U --> W[Log, update seen_hashes]
	V --> W
	W --> X[Increment saved_count]
	X --> N
	N --> I
	I --> Y[Log matched messages, saved attachments]
	Y --> Z{--open-folder?}
	Z -- Yes --> AA[Open output folder]
	Z -- No --> AB[Log completion]
	AA --> AB
	AB --> AC([End])
	E --> AC
```
