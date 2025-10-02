import os
import shutil
import zipfile
import argparse
from pathlib import Path


class XLSXManager:
    def __init__(
        self,
        workbooks_dir,
        xmls_dir,
    ):
        self.workbooks_dir = Path(workbooks_dir)
        self.xmls_dir = Path(xmls_dir)

        # Ensure directories exist
        self.workbooks_dir.mkdir(parents=True, exist_ok=True)
        self.xmls_dir.mkdir(parents=True, exist_ok=True)

    def unpack_xlsx(self, xlsx_filename):
        """
        Unpacks an XLSX file into XML components.

        Args:
            xlsx_filename (str): Name of the XLSX file in the workbooks
                               directory

        Returns:
            Path: Path to the directory containing the unpacked XML files
        """
        xlsx_path = self.workbooks_dir / xlsx_filename
        if not xlsx_path.exists():
            raise FileNotFoundError(f"XLSX file not found: {xlsx_path}")

        # Get the name without extension for the output directory
        name_without_ext = xlsx_path.stem
        output_dir = self.xmls_dir / name_without_ext

        # Remove existing directory if it exists
        if output_dir.exists():
            shutil.rmtree(output_dir)

        # Create the output directory
        output_dir.mkdir(parents=True)

        # Unzip the XLSX file
        with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)

        return output_dir

    def pack_xlsx(self, dir_name, output_filename=None):
        """
        Packs an unzipped XLSX directory back into an XLSX file.

        Args:
            dir_name (str): Name of the directory in xmls folder containing
                          the XLSX components
            output_filename (str, optional): Name for the output XLSX file.
                                         If None, uses the directory name.

        Returns:
            Path: Path to the created XLSX file
        """
        source_dir = self.xmls_dir / dir_name
        if not source_dir.exists():
            msg = f"Source directory not found: {source_dir}"
            raise FileNotFoundError(msg)

        # If no output filename is provided, use the directory name
        if output_filename is None:
            output_filename = f"{dir_name}.xlsx"
        elif not output_filename.endswith('.xlsx'):
            output_filename = f"{output_filename}.xlsx"

        output_path = self.workbooks_dir / output_filename

        # Create the ZIP file
        zip_params = {'mode': 'w', 'compression': zipfile.ZIP_DEFLATED}
        with zipfile.ZipFile(output_path, **zip_params) as zip_file:
            # Walk through all files in the directory
            for root, _, files in os.walk(source_dir):
                for file in files:
                    file_path = Path(root) / file
                    # Calculate the archive path (relative to source_dir)
                    archive_path = file_path.relative_to(source_dir)
                    zip_file.write(file_path, archive_path)

        return output_path


def parse_args():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='XLSX file unpacker and packer utility'
    )

    parser.add_argument(
        'action',
        choices=['pack', 'unpack'],
        help='Action to perform: pack or unpack XLSX file'
    )

    parser.add_argument(
        'name',
        help='For unpack: name of the XLSX file in workbooks directory. '
             'For pack: name of the directory in xmls directory'
    )

    parser.add_argument(
        '-o', '--output',
        help='Output name (optional). For pack: output XLSX filename. '
             'For unpack: output directory name'
    )

    parser.add_argument(
        '--workbooks-dir',
        default='assets/workbooks',
        help='Directory containing XLSX files (default: assets/workbooks)'
    )

    parser.add_argument(
        '--xmls-dir',
        default='assets/xmls',
        help='Directory for XML files (default: assets/xmls)'
    )

    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    try:
        # Create an instance of XLSXManager with configured directories
        manager = XLSXManager(args.workbooks_dir, args.xmls_dir)

        if args.action == "unpack":
            # Ensure .xlsx extension
            xlsx_name = args.name
            if not xlsx_name.endswith('.xlsx'):
                xlsx_name += '.xlsx'

            result = manager.unpack_xlsx(xlsx_name)
            print(f"Successfully unpacked to: {result}")

        else:  # pack
            result = manager.pack_xlsx(args.name, args.output)
            print(f"Successfully packed to: {result}")

    except FileNotFoundError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
