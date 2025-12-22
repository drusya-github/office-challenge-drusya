# Challenge: Document Redaction with Tracking & Confidentiality Header

Create a button that will redact sensitive information from a Word document, add a confidentiality header, and enable the Tracking Changes feature to log these modifications.

## Requirements:

1. **Redact Sensitive Information**
    - Retrieve the document's complete content
    - Locate and identify sensitive information (emails, phone numbers, social security numbers)
    - Replace this information with redaction markers in the document
2. **Add Confidential Header**
    - Insert a header at the top of the document stating "CONFIDENTIAL DOCUMENT"
    - Ensure this header addition is tracked by the Tracking Changes feature
3. **Enable Tracking Changes**
    - Use the Office Tracking Changes API to enable tracking changes
    - Make sure to only use Tracking Changes if the Word API is available
    [Word JavaScript API requirement set 1.5 - Office Add-ins | Microsoft Learn](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-1-5-requirement-set?view=word-js-preview)

## Technical Requirements

- Code must be with TypeScript.
- You are free to use any framework or build tool (Vite, Next.js, etc.),  DON'T use the minimum setup of this repo.
- Use self written CSS for styling instead of external libraries, we expect good design and craftsmanship.
- The solution must run in Word on the web or Word desktop.
- Don't create a public forked repository, otherwise your solution will be disqualified. Share your solution as a private repository or a zip file.
- One of the evaluation criteria will be code quality, so please ensure your code is clean, well-structured, and follows best practices and crasftsmanship.


## Testing Your Solution
Use the attached Document-To-Be-Redacted.docx file to test your solution. The document contains various instances of sensitive information that should be redacted when your add-in is executed.

We will use a different document to evaluate your solution, so ensure that your redaction logic is robust and can handle various scenarios.


## Run the Challenge

1. `npm install`
2. `npm start`
   - Starts local server on port 3000.
   - Compiles TypeScript.
   - Attempts to sideload to Word.

If automatic sideloading fails, please [sideload the manifest manually](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).


## Submission

1. Ensure your solution meets all the requirements outlined above.
2. Share your solution as a zip file (without the node_modules folder).
3. Include any necessary instructions to run and test your solution, BUT it should be straightforward to run following the steps in the "Run the Challenge" section.
4. Submit your solution before the deadline specified in the challenge announcement.

Good luck, and we look forward to seeing your innovative solutions!

