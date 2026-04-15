# **App Name**: Formulytics

## Core Features:

- User Authentication: Secure user registration and login with email and password, ensuring access to personalized data.
- Template Upload & Metadata Extraction: Allow users to upload .xlsx template files, automatically extracting and storing metadata (sheets, columns) in Firestore, while saving the original file in Firebase Storage. This uses client-side SheetJS for extraction.
- Template Management Dashboard: A personal dashboard listing all user-specific templates from Firestore, enabling users to apply, delete, or download their original templates.
- Raw Data Upload & Preview: Provide a simple interface for uploading raw .xlsx data files, including a real-time preview of the content before processing. Uses client-side SheetJS for reading.
- Formula Application & Treated File Download: Client-side application of formulas from a selected template to a raw data file, generating a new .xlsx file ready for download. This ensures raw data privacy as files are not stored on Firebase. Uses SheetJS for processing.
- AI Formula Assistant Tool: A generative AI tool that assists users by suggesting or explaining Excel formulas based on natural language descriptions, helping to refine their templates or troubleshoot issues.

## Style Guidelines:

- Primary color: Deep violet (#3D22C3) to convey professionalism and sophistication in data handling. It contrasts well with the light background for clear UI elements.
- Background color: A very light, subtle cool-gray with a hint of violet (#F7F6F9), maintaining a clean and calm workspace without overwhelming the eye.
- Accent color: A vibrant medium blue (#477EEB), analogous to the primary, used for highlighting interactive elements, active states, and calls to action, providing clear visual guidance.
- Headline and body font: 'Inter' (sans-serif) for its modern, clean, and objective aesthetic, ensuring excellent readability for data and technical information.
- Utilize modern, clean line-style or solid vector icons throughout the application to maintain a polished and intuitive user interface, especially for file management and actions.
- Implement a classic 'sidebar + main content' layout, where a dark navigation sidebar provides easy access to sections and settings, while the primary workspace occupies the main content area for focused tasks.
- Incorporate subtle, functional animations for feedback during file uploads, processing progress, and state changes, enhancing the user experience without distractions.