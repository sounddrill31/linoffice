#!/bin/bash

# Check if setup.sh exists and extract USER_APPLICATIONS_DIR and APPDATA_PATH
if [[ ! -f "setup.sh" ]]; then
  echo "Error: setup.sh not found in the current directory."
  exit 1
fi

USER_APPLICATIONS_DIR=$(grep -E '^\s*USER_APPLICATIONS_DIR=' setup.sh | cut -d'=' -f2 | tr -d '"' | tr -d "'")
if [[ -z "$USER_APPLICATIONS_DIR" ]]; then
  echo "Error: USER_APPLICATIONS_DIR not found in setup.sh. The uninstallation is incomplete."
fi

APPDATA_PATH=$(grep -E '^\s*APPDATA_PATH=' setup.sh | cut -d'=' -f2 | tr -d '"' | tr -d "'")
if [[ -z "$APPDATA_PATH" ]]; then
  echo "Error: APPDATA_PATH not found in setup.sh. The uninstallation is incomplete."
fi

# Check if APPDATA_PATH exists and delete if it does
if [[ -d "$APPDATA_PATH" ]]; then
  rm -r "$APPDATA_PATH"
else
  echo "Error: Directory $APPDATA_PATH does not exist. The uninstallation is incomplete."
fi

# Check if USER_APPLICATIONS_DIR exists
if [[ ! -d "$USER_APPLICATIONS_DIR" ]]; then
  echo "Error: Directory $USER_APPLICATIONS_DIR does not exist. The uninstallation is incomplete."
fi

# Find .desktop files containing linoffice.sh in Exec= line
DESKTOP_FILES=$(find "$USER_APPLICATIONS_DIR" -type f -name "*.desktop" -exec grep -l "Exec=.*linoffice.sh" {} \;)
if [[ -n "$DESKTOP_FILES" ]]; then
  echo "The following .desktop files will be deleted:"
  echo "$DESKTOP_FILES"
  read -p "Do you want to proceed with deletion? (y/n): " confirm
  if [[ "$confirm" == "y" || "$confirm" == "Y" ]]; then
    while IFS= read -r file; do
      rm -f "$file" && echo "Deleted: $file"
    done <<< "$DESKTOP_FILES"
  else
    echo "Deletion of .desktop files aborted."
  fi
else
  echo "No .desktop files created by LinOffice were found."
fi

# Find all files and folders in the same directory as uninstall.sh (excluding itself)
CURRENT_DIR="$(dirname "$(realpath "$0")")"
SCRIPT_NAME="$(basename "$0")"

FILES_TO_DELETE=$(find "$CURRENT_DIR" -maxdepth 1 -not -name "uninstall.sh")
if [[ -n "$FILES_TO_DELETE" ]]; then
  echo "The following files and folders will be deleted recursively:"
  echo "$FILES_TO_DELETE"
  read -p "Do you want to proceed with deletion? (y/n): " confirm
  if [[ "$confirm" == "y" || "$confirm" == "Y" ]]; then
    find "$CURRENT_DIR" -maxdepth 1 -not -name "$SCRIPT_NAME" -exec rm -rf {} \;
    echo "Files and folders deleted."
  else
    echo "Deletion of files and folders aborted."
  fi
else
  echo "No files or folders to delete in $CURRENT_DIR."
fi

# Ask to delete the Windows container and its data
read -p "Do you want to delete the Windows container and all its data as well? (y/n): " confirm
if [[ "$confirm" == "y" || "$confirm" == "Y" ]]; then
  if ! command -v podman &> /dev/null; then
    echo "Error: Podman is not installed or not accessible."
  else
    # Stop and remove the LinOffice container
    if ! podman stop LinOffice &> /dev/null; then
      echo "Error: Could not stop the LinOffice container. It may not exist."
    else
      echo "Stopped LinOffice container."
    fi

    if ! podman rm -f LinOffice &> /dev/null; then
      echo "Error: Could not delete the LinOffice container."
    else
      echo "Deleted LinOffice container."
    fi

    # Remove the linoffice_data volume
    if ! podman volume rm linoffice_data &> /dev/null; then
      echo "Error: Could not delete the linoffice_data volume."
    else
      echo "Deleted linoffice_data volume."
    fi
  fi
else
  echo "Windows container and data deletion aborted."
fi

# Delete the uninstall.sh script itself
echo "Deleting the uninstall script itself."
rm -f "$0"
echo "Script deleted."
