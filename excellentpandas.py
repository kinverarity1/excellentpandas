import pandas as pd
import xlwings as xw

__version__ = '0.1.1'


def show_in_excel(
    df, return_df=False, activate_excel=True, print_info=False, max_cols=None
):
    """Show DataFrame/Series in Excel."""
    if print_info:
        if isinstance(df, pd.DataFrame):
            print(df.info(max_cols=max_cols))
        else:
            print(df.info())
    if not isinstance(df, pd.DataFrame):
        tmp_df = df.to_frame()
    else:
        tmp_df = df
    book = xw.Book()
    book.sheets[0].range("A1").value = tmp_df
    if activate_excel:
        book.activate(steal_focus=True)
    if return_df:
        return df


def via_excel(df):
    """Show DataFrame/Series in Excel and return the DataFrame/Series.

    This can be added to existing scripts without
    any change to the script's behaviour:

        >>> from excellentpandas import show_in_excel_pipe
        >>> result = df.do_something().pipe(show_in_excel_pipe)

    """
    return show_in_excel(df, return_df=True)


def via_info_excel(df, max_cols=None):
    """Show DataFrame/Series in Excel and return the DataFrame/Series.

    This can be added to existing scripts without
    any change to the script's behaviour:

        >>> from excellentpandas import show_in_excel_pipe
        >>> result = df.do_something().pipe(show_in_excel_pipe)

    """
    return show_in_excel(df, return_df=True, print_info=True, max_cols=max_cols)
