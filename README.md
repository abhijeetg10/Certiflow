# Certiflow 🏆

A Full-Stack Web Application for Bulk Certificate Generation and Automated Email Sending.

## Features

- **Upload CSV Data**: Maps student names to their emails.
- **Certificate Template Options**: Attach a PDF or PNG certificate template.
- **Dynamic PDF Generation**: Automates personalized certificate generation centrally placing the names.
- **Gmail Integration**: Connects directly to Gmail to distribute certificates via email securely.
- **ZIP Download Archive**: Creates a downloadable ZIP folder of all generated certificates.
- **Interactive Interface**: Responsive dark-mode UI with a real-time progress bar.

## Prerequisites

- Node.js installed on your machine.
- A Gmail account with 2-Step Verification enabled.

## Installation

1. Clone or download this project, then navigate inside the directory:

   ```bash
   cd certificate-automation
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

## Running the App

1. Start the server:

   ```bash
   node server.js
   ```

2. Open your browser and go to: **[http://localhost:3000](http://localhost:3000)**

## How to use Gmail App Password

Since standard passwords are blocked by Google for external app SMTP usage, you need to use an App Password:

1. Go to your **Google Account Settings** > **Security**.
2. Enable **2-Step Verification** if you haven't already.
3. Search for **App passwords** in your Google Account security search bar.
4. Create a new App password (you can name it "Certiflow").
5. Copy the generated 16-character password and paste it into the "App Password" field on the website. **Note: No spaces are required!**

## Folder Structure

- `/public/`: Contains the frontend HTML, CSS, and Vanilla JavaScript.
- `/uploads/`: Temporary directory used to store CSVs and templates before processing. Automatically cleaned up after execution.
- `/generated/`: Temporary directory holding completed PDF certificates and ZIP files. Automatically cleaned up.
- `server.js`: The Express backend handling routing, rendering, PDF parsing, and email dispatch.
- `package.json`: Contains project dependencies and scripts.

## Advanced Usage

- **Body Customization**: Include `[Name]` in the Email Body field on the UI, and the system will automatically replace it with the student's name from the corresponding row in the CSV file.
- **Text Placement**: Leaving `X Position` and `Y Position` blank gracefully auto-centers the text. Setting specific values accurately moves the text.
