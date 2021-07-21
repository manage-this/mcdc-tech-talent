import re
from pathlib import Path

import numpy as np
import pandas as pd


def load_msa_lookup(data_path):
    """ Loads a lookup table for various MSA specific attributes used
    throughout the project.

    :param data_path: a pathlib.Path() object, location of data file folder

    :return msa_lookup: a dictionary containing MSA code to peer type lookup

    """

    lookup_dir = 'lookups'
    msa_file_name = 'lk_msa.xlsx'

    lookup_file_path = data_path / lookup_dir / msa_file_name

    try:
        lk_msa = pd.read_excel(lookup_file_path)
        msa_lookup = dict(zip(lk_msa['area'], lk_msa['peer_type']))
        print("...MSA lookup table loaded.")
        return msa_lookup
    except Exception as e:
        print("MSA lookup table loading failed - {}".format(e))
        return None


def load_oes_lookup(data_path):
    """ Loads a crosswalk for SOC codes provided by Bureau of Labor Statistics
    https://www.bls.gov/oes/soc_2018.htm

    :param data_path: A pathlib.Path() object, location of data file folder

    :return soc_lookup: A dictionary containing crosswalks for various years
    """

    lookup_dir = 'lookups'
    oes_file_name = 'oes_2019_hybrid_structure.xlsx'

    lookup_file_path = data_path / lookup_dir / oes_file_name

    try:
        # skip first 5 rows that contain notes
        df = pd.read_excel(lookup_file_path, sheet_name=0, skiprows=5)

        # rename column headers for easier manipulation
        df.columns = ['oes_code_2019', 'oes_title_2019',
                      'soc_code_2018', 'soc_title_2018',
                      'oes_code_2018', 'oes_title_2018',
                      'soc_code_2010', 'soc_title_2010', 'notes']

        # create SOC2010 to OES2019 for 2015 and 2016
        # create OES2018 to OES2019 for unmatched 2015 and 2016, and 2017-2020
        soc1019_dict = dict(zip(df['soc_code_2010'], df['oes_code_2019']))
        oes1819_dict = dict(zip(df['oes_code_2018'], df['oes_code_2019']))
        oes1919_dict = dict(zip(df['oes_code_2019'], df['oes_code_2019']))
        oes19_dict = dict(zip(df['oes_code_2019'], df['oes_title_2019']))

        soc_lookup = {'1019': soc1019_dict,
                      '1819': oes1819_dict,
                      '1919': oes1919_dict,
                      '19': oes19_dict}
        print("...SOC lookup tables loaded.")
        return soc_lookup
    except Exception as e:
        print("SOC lookup table loading failed - {}".format(e))
        return None


def much_consistency(df):
    """ Helper function to handle known inconsistencies in raw data files

    :param df: A pandas dataframe, BLS OEWS data set

    :return df: A pandas dataframe
    """

    # Force lower case column headers
    df.columns = map(str.lower, df.columns)
    print("...Forced column headers to lower case.")

    # Rename column headers used inconsistently from year to year
    df.rename(columns={'occ_group': 'o_group',
                       'loc quotient': 'loc_quotient',
                       'area_name': 'area_title'}, inplace=True)
    print("...Standardized column header names.")

    return df


def map_soc(df, report_year, soc_lookup):
    """ Helper function to apply SOC code crosswalks for older BLS OEWS data.
    See: https://www.bls.gov/oes/soc_2018.htm

    :param df: A pandas dataframe, BLS OEWS data set
    :param report_year: An integer, release year of data set
    :param soc_lookup: A dictionary, crosswalk loaded via load_oes_lookup()

    :return df: A pandas dataframe
    """

    soc_mappings = soc_lookup['1919']

    if report_year in [2014, 2015, 2016]:
        soc_mappings = soc_lookup['1019']
    elif report_year in [2017, 2018]:
        soc_mappings = soc_lookup['1819']
    elif report_year >= 2019:
        soc_mappings = soc_lookup['1919']

    # Apply SOC code crosswalk
    df['oes_code_2019'] = df[df['o_group'] == 'detailed']['occ_code'].map(soc_mappings)

    # Map the couple codes that didn't have a 2010 to 2019 conversion separately
    df_matched = df[(~df['oes_code_2019'].isnull()) & (df['o_group'] == 'detailed')].copy()
    df_missing = df[(df['oes_code_2019'].isnull()) & (df['o_group'] == 'detailed')].copy()
    df_missing['oes_code_2019'] = df_missing['occ_code'].map(soc_lookup['1819'])

    # Map the total and major rows separately, because these aren't in the federal crosswalk
    df_totals = df[df['o_group'] != 'detailed'].copy()
    df_totals['oes_code_2019'] = df_totals['occ_code']
    df_totals['oes_title_2019'] = df_totals['occ_title']

    # Recombine all rows
    df_combined = pd.concat([df_matched, df_missing])
    df_combined['oes_title_2019'] = df_combined['oes_code_2019'].map(soc_lookup['1919'])

    df_f = pd.concat([df_combined, df_totals])

    print("...Applied SOC crosswalks.")

    return df_f


def map_peer_type(df, msa_lookup):
    """ Helper function to apply peer group mapping to MSA locations

    :param df: A pandas dataframe, BLS OEWS data set
    :param msa_lookup: a dictionary containing MSA code to peer type lookup

    :return df: A pandas dataframe

    """

    # Map MSA areas to peer type categories
    df['peer_type'] = df['area'].map(msa_lookup)
    df['peer_type'].fillna("All Other MSA", inplace=True)
    print("...Applied MSA lookups.")

    return df


def map_nulls(df):
    """ Helper function to convert missing or unreported values into nulls

    :param df: A pandas dataframe, BLS OEWS data set

    :return df: a pandas dataframe
    """
    null_replace = {'#': np.nan, '*': np.nan, '**': np.nan}
    df.replace(null_replace, inplace=True)
    print("...Replaced missing and unreported values with nulls.")

    return df


def clean_bls(data_path):
    """

    :param data_path: A pathlib.Path() object, location of data file folder

    :return df_merged: A pandas dataframe, cleaned and merged dataframe
                       containing all available years of BLS OEWS data
    """
    df_merged = pd.DataFrame()

    bls_folder = data_path / 'bls'

    soc_lookup = load_oes_lookup(data_path)
    msa_lookup = load_msa_lookup(data_path)

    for f in bls_folder.glob('**/*.xlsx'):
        print('Processing {} ...'.format(f))
        df = pd.read_excel(f)
        df = much_consistency(df)

        report_year = int(re.search(r"(M)([1-9]\d{3,})(_)", str(f))[2])

        df = map_soc(df, report_year, soc_lookup)
        df = map_peer_type(df, msa_lookup)
        df = map_nulls(df)
        df['report_year'] = report_year

        df = df[['report_year', 'area', 'area_title', 'peer_type',
                 'oes_code_2019', 'oes_title_2019', 'o_group',
                 'tot_emp', 'emp_prse', 'jobs_1000', 'loc_quotient',
                 'h_mean', 'a_mean', 'mean_prse',
                 'h_pct10', 'h_pct25', 'h_median', 'h_pct75', 'h_pct90',
                 'a_pct10', 'a_pct25', 'a_median', 'a_pct75', 'a_pct90']
                ].reset_index(drop=True)

        df_merged = pd.concat([df_merged, df])
        print('...{} processing complete.'.format(f))

    df_merged = df_merged[df_merged['peer_type'] != 'All Other MSA']
    #df_merged = df_merged[df_merged['oes_code_2019'].str.contains('15-')]

    return df_merged


def split_bls_ogroup(df):
    """ Helper function to split the BLS OEWS data into three separate
    dataframes each containing a different o_group level.

    :param df: A pandas dataframe, BLS OEWS data that has been cleaned

    :return bls_total: A pandas dataframe, o_group = total
    :return bls_major: A pandas dataframe, o_group = major
    :return bls_detailed: A pandas dataframe, o_group = detailed
    """
    bls_total = df[df['o_group'] == 'total']
    bls_major = df[df['o_group'] == 'major']
    bls_detailed = df[df['o_group'] == 'detailed']

    print("BLS data split into total, major, detailed.")
    return bls_total, bls_major, bls_detailed


def process_bls(data_path, db_path):
    """ Perform data cleaning processes on BLS data and save results

    :param data_path: A pathlib.Path() object, location of project's data dir
    :param db_path: A pathlib.Path() object, location of project's db dir

    :return: None
    """
    bls_merged = clean_bls(data_path)
    bls_total, bls_major, bls_detailed = split_bls_ogroup(bls_merged)

    print("Saving BLS files to {}".format(db_path))
    bls_total.to_excel(db_path / 'bls.total.xlsx', index=False)
    bls_major.to_excel(db_path / 'bls.major.xlsx', index=False)
    bls_detailed.to_excel(db_path / 'bls.detailed.xlsx', index=False)


if __name__ == "__main__":
    DATA_DIR = 'data'
    DB_DIR = 'db'
    DATA_PATH = Path() / DATA_DIR
    DB_PATH = Path() / DB_DIR

    process_bls(DATA_PATH, DB_PATH)
