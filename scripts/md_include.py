from pathlib import Path
import re
from loguru import logger

# Map file extensions to their corresponding code block languages
EXTENSION_MAP = {
    ".py": "python",
    ".rs": "rust",
    ".java": "java",
    ".js": "javascript",
    ".html": "html",
    ".css": "css",
    ".sh": "bash",
    # Add more extensions and languages as needed
}

def md_include(source_md_file_path, target_md_file_path):
    # Convert the input path to a Path object
    source_md_file_path = Path(source_md_file_path)
    target_md_file_path = Path(target_md_file_path)

    # Regular expression to find !md_include directives
    include_pattern = re.compile(r'\s*!md_include\s*["\']([^"\']+)["\']')

    # Read the markdown file
    with source_md_file_path.open('r', encoding='utf-8') as md_file:
        md_content = md_file.readlines()

    # Process each line, looking for !md_include directive
    new_md_content = []
    for line in md_content:
        # Find all !md_include directives in the line
        matches = include_pattern.findall(line)

        if matches:
            for included_file_path in matches:
                # Make the path relative to the markdown file
                full_included_path = source_md_file_path.parent / included_file_path
                file_extension = full_included_path.suffix.lower()  # Get the file extension

                # Determine the code block language based on file extension
                language = EXTENSION_MAP.get(file_extension, "")  # Default to empty if not found

                # Read the content of the file to be included
                try:
                    with full_included_path.open('r', encoding='utf-8') as included_file:
                        included_content = included_file.read()

                        # Replace the !md_include directive with the content of the included file
                        # Wrap it in a code block with the appropriate language
                        if language:
                            line = line.replace(
                                f'!md_include "{included_file_path}"', 
                                f"```{language}\n{included_content}\n```"
                            )
                        else:
                            # No language specified, default to plain code block
                            line = line.replace(
                                f'!md_include "{included_file_path}"', 
                                f"```\n{included_content}\n```"
                            )
                except FileNotFoundError:
                    logger.error(f"Error: The file '{full_included_path}' could not be found.")
                    line = line.replace(f'!md_include "{included_file_path}"', f"**Error: Could not include file '{full_included_path}'**")

        new_md_content.append(line)

    # Write the updated content to a new file or overwrite the original
    with target_md_file_path.open('w', encoding='utf-8') as new_md_file:
        new_md_file.writelines(new_md_content)

    logger.info(f"Processed markdown file saved as: {target_md_file_path}")

# Example usage
if __name__ == "__main__":
    # Replace 'example.md' with the path to your actual markdown file
    md_include('docs/readme.md', 'README.md')