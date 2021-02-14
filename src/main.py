from utilities import *
import logging
import os, sys
from time import time
import argparse
import helpers
from helpers import TerritoryFinder, OutletAllocation

def main():

    try:
        t0 = time()
        # Set up logging
        logger = logging.getLogger('sales_support_helper_application')
        logger.setLevel(logging.DEBUG)
        # create file handler which logs even debug messages
        os.makedirs('logs', exist_ok=True)
        fh = logging.FileHandler(u"./logs/sales-support-helper.log", "w")
        fh.setLevel(logging.DEBUG)
        # create console handler with a higher log level
        ch = logging.StreamHandler(sys.stdout)
        ch.setLevel(logging.INFO)
        # create formatter and add it to the handlers
        formatter = logging.Formatter(fmt="%(asctime)s %(levelname)s: %(message)s", datefmt="%Y-%m-%d - %H:%M:%S")
        fh.setFormatter(formatter)
        ch.setFormatter(formatter)
        # add the handlers to the logger
        logger.addHandler(fh)
        logger.addHandler(ch)

        logger.debug("Parsing command line")

        parser = argparse.ArgumentParser(description="Sales support helper (Outlet Allocation and Territory Finder operations)")

        # Parse arguments
        parser.add_argument("-p", "--operation", type=str, action='store',
            help="Please specify the operation to execute", required=False, default='None')
        parser.add_argument("-c", "--coord_file", type=str, action='store',
            help="Please specify the file with coordinates", required=False, default='coordinates.xlsx')
        parser.add_argument("-r", "--report_file", type=str, action='store',
            help="Please specify the 'Report Territory Management' file", required=False, default='report.xlsx')
        parser.add_argument("-o", "--output_file", type=str, action='store',
            help="Please specify the output file name", required=False)
        parser.add_argument("-s", "--samples_threshold", type=int, action='store',
            help="Specify a minimal samples number in a class", default=2, required=False)

        args = parser.parse_args()

        if args.operation != 'None':
            operation = args.operation
            coord_file = args.coord_file
            report_file = args.report_file
            output_file = os.path.splitext(report_file)[0] + " updated.xlsx" if args.output_file == None else args.output_file
            if output_file == report_file:
                logger.error(f"Input and output reports have the identical name '{output_file}' " \
                    "Please set a unique name for the output report (--output_file parameter).")
                sys.exit(1)
            samples_threshold = args.samples_threshold
        else:
            logger.debug("No command line args found. Scanning the data folder...")
            operation, coord_file, report_file = select_input_files()
            output_file = os.path.splitext(report_file)[0] + " updated.xlsx"
            samples_threshold = 3
        
        # Initialize
        if operation.lower() == 'territoryfinder':
            hlpr = TerritoryFinder(coord_file, report_file, output_file, samples_threshold)
        elif operation.lower() == 'outletallocation':
            hlpr = OutletAllocation(coord_file, report_file, output_file, samples_threshold)
        else:
            logger.error(f"Operation type '{operation}' is not supported.")
            sys.exit(1)
            
        total_steps = 4

        logger.info(f"Step 1 of {total_steps}: Loading and prepare data")
        hlpr.load_data()

        logger.info(f"Step 2 of {total_steps}: Validate the model")
        hlpr.validate()

        logger.info(f"Step 3 of {total_steps}: Prepare report")
        hlpr.get_report()

        logger.info(f"Step 4 of {total_steps}: Save report")
        hlpr.save_report()

        logger.info(f"Total elapsed time {time() - t0:.3f} sec.")
    
    except Exception as err:
        logger.exception(err)

    finally:
        pass
        sys.exit(1)

if __name__ == '__main__':
    main()