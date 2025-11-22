import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, date, timedelta
import pytz

# Configure the page
st.set_page_config(
    page_title="HKU Course Explorer",
    page_icon="📚",
    layout="wide"
)

# Hong Kong timezone
HK_TIMEZONE = pytz.timezone('Asia/Hong_Kong')


@st.cache_data
def load_and_preprocess_data():
    """Load and preprocess the Excel data"""
    try:
        # Load the Excel file
        df = pd.read_excel('2025-26_class_timetable_20250831.xlsx')

        # Strip whitespace from column names
        df.columns = df.columns.str.strip()

        # Convert all columns to string to handle mixed types
        for col in df.columns:
            df[col] = df[col].astype(str)
            # Replace 'nan' with empty string
            df[col] = df[col].replace('nan', '')

        # Clean and transform OFFER DEPT column
        if 'OFFER DEPT' in df.columns:
            # Replace "Business and Econom" with "Business & Economics"
            df['OFFER DEPT'] = df['OFFER DEPT'].str.replace(
                'Business and Econom', 'Business & Economics', regex=False)

            # Replace "Faculty of ..." with "... Faculty"
            df['OFFER DEPT'] = df['OFFER DEPT'].str.replace(
                'Faculty of ', '', regex=False)
            df['OFFER DEPT'] = df['OFFER DEPT'].apply(
                lambda x: x + ' Faculty' if x and not x.endswith('Faculty') else x)

            # Replace "and" with "&"
            df['OFFER DEPT'] = df['OFFER DEPT'].str.replace(
                ' and ', ' & ', regex=False)

            # Replace "Department" with "Dept"
            df['OFFER DEPT'] = df['OFFER DEPT'].str.replace(
                'Department', 'Dept', regex=False)

        # Add whitespace after each comma in instructor names
        if 'INSTRUCTOR' in df.columns:
            df['INSTRUCTOR'] = df['INSTRUCTOR'].str.replace(
                ',', ', ', regex=False)
            # Remove any double spaces that might have been created
            df['INSTRUCTOR'] = df['INSTRUCTOR'].str.replace(
                '  ', ' ', regex=False)

        # Create a WEEKDAY column by combining Mon, Tue, Wed, etc.
        weekday_columns = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN']

        # Create WEEKDAY column - find which day has a value
        def get_weekday(row):
            for day in weekday_columns:
                if day in row and row[day] != '' and pd.notna(row[day]) and row[day] != 'nan':
                    return day
            return 'Unknown'

        df['WEEKDAY'] = df.apply(get_weekday, axis=1)

        # Convert date columns to datetime
        date_columns = ['START DATE', 'END DATE']
        for col in date_columns:
            if col in df.columns:
                # Handle empty strings and convert to datetime
                df[col] = df[col].replace('', pd.NaT)
                df[col] = pd.to_datetime(df[col], errors='coerce')

        # Handle time columns
        time_columns = ['START TIME', 'END TIME']
        for col in time_columns:
            if col in df.columns:
                # Replace empty strings with NaT
                df[col] = df[col].replace('', pd.NaT)
                # Convert to datetime to extract time
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.time

        return df

    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()


def get_hk_time():
    """Get current time in Hong Kong timezone"""
    now_utc = datetime.now(pytz.utc)
    now_hk = now_utc.astimezone(HK_TIMEZONE)
    return now_hk


def create_filters_panel(df):
    """Create dynamic filter panel based on available columns"""
    st.sidebar.header("🔍 Filters & Sorting")

    filters = {}

    # Course Code filter
    if 'COURSE CODE' in df.columns:
        # Get unique course codes, filter out empty strings and NaN, then sort
        course_codes = df['COURSE CODE'].unique()
        course_codes = [code for code in course_codes if code !=
                        '' and pd.notna(code) and code != 'nan']
        course_codes = ['All'] + sorted(course_codes)
        selected_course = st.sidebar.selectbox(
            "Course Code",
            course_codes,
            key="course_code_filter"
        )
        if selected_course != 'All':
            filters['COURSE CODE'] = selected_course

    # Class Section filter
    if 'CLASS SECTION' in df.columns:
        # Get unique class sections, filter out empty strings and NaN, then sort
        class_sections = df['CLASS SECTION'].unique()
        class_sections = [section for section in class_sections if section != '' and pd.notna(
            section) and section != 'nan']
        class_sections = ['All'] + sorted(class_sections)
        selected_section = st.sidebar.selectbox(
            "Class Section",
            class_sections,
            key="class_section_filter"
        )
        if selected_section != 'All':
            filters['CLASS SECTION'] = selected_section

    # Curriculum filter (renamed from ACAD_CAREER)
    if 'ACAD_CAREER' in df.columns:
        curricula = df['ACAD_CAREER'].unique()
        curricula = [cur for cur in curricula if cur !=
                     '' and pd.notna(cur) and cur != 'nan']
        curricula = ['All'] + sorted(curricula)
        selected_curriculum = st.sidebar.selectbox(
            "Curriculum",
            curricula,
            key="curriculum_filter"
        )
        if selected_curriculum != 'All':
            filters['ACAD_CAREER'] = selected_curriculum

    # Department filter (OFFER DEPT)
    if 'OFFER DEPT' in df.columns:
        departments = df['OFFER DEPT'].unique()
        departments = [dept for dept in departments if dept !=
                       '' and pd.notna(dept) and dept != 'nan']
        departments = ['All'] + sorted(departments)
        selected_department = st.sidebar.selectbox(
            "Department",
            departments,
            key="department_filter"
        )
        if selected_department != 'All':
            filters['OFFER DEPT'] = selected_department

    # Weekday filter
    if 'WEEKDAY' in df.columns:
        weekdays = df['WEEKDAY'].unique()
        weekdays = [day for day in weekdays if day !=
                    '' and pd.notna(day) and day != 'nan']
        weekdays = ['All'] + sorted(weekdays)
        selected_weekday = st.sidebar.selectbox(
            "Weekday",
            weekdays,
            key="weekday_filter"
        )
        if selected_weekday != 'All':
            filters['WEEKDAY'] = selected_weekday

    # Venue filter
    if 'VENUE' in df.columns:
        venues = df['VENUE'].unique()
        venues = [v for v in venues if v != '' and pd.notna(v) and v != 'nan']
        venues = ['All'] + sorted(venues)
        selected_venue = st.sidebar.selectbox(
            "Venue",
            venues,
            key="venue_filter"
        )
        if selected_venue != 'All':
            filters['VENUE'] = selected_venue

    # Venue is provided checkbox
    st.sidebar.markdown("---")
    venue_provided = st.sidebar.checkbox(
        "Venue is provided", value=True, key="venue_provided_filter")
    filters['venue_provided'] = venue_provided

    # Time range filters
    st.sidebar.markdown("---")
    st.sidebar.subheader("Time Range Filters")

    if 'START TIME' in df.columns or 'END TIME' in df.columns:
        # Get available time values for the time inputs
        all_times = []
        if 'START TIME' in df.columns:
            start_times = [t for t in df['START TIME'].dropna() if t != '']
            all_times.extend(start_times)
        if 'END TIME' in df.columns:
            end_times = [t for t in df['END TIME'].dropna() if t != '']
            all_times.extend(end_times)

        if all_times:
            # Find min and max times for the range
            min_time = min(all_times)
            max_time = max(all_times)

            # Create time options for the selectors
            time_options = [
                time(8, 0), time(9, 0), time(10, 0), time(11, 0), time(12, 0),
                time(13, 0), time(14, 0), time(
                    15, 0), time(16, 0), time(17, 0),
                time(18, 0), time(19, 0), time(20, 0)
            ]

            # Filter time options to be within data range
            time_options = [
                t for t in time_options if min_time <= t <= max_time]

            col1, col2 = st.sidebar.columns(2)
            with col1:
                start_time_filter = st.selectbox(
                    "Start Time From",
                    options=time_options,
                    index=0,
                    format_func=lambda x: x.strftime('%H:%M'),
                    key="start_time_from_filter"
                )
            with col2:
                end_time_filter = st.selectbox(
                    "End Time To",
                    options=time_options,
                    index=len(time_options)-1,
                    format_func=lambda x: x.strftime('%H:%M'),
                    key="end_time_to_filter"
                )

            filters['time_range'] = (start_time_filter, end_time_filter)

    # Date range filter
    if 'START DATE' in df.columns and not df['START DATE'].isna().all():
        # Filter out NaT values for date calculation
        valid_dates = df[df['START DATE'].notna()]
        if not valid_dates.empty:
            min_date = valid_dates['START DATE'].min().date()
            max_date = valid_dates['END DATE'].max().date()

            st.sidebar.markdown("---")
            st.sidebar.subheader("Date Range")
            col1, col2 = st.sidebar.columns(2)
            with col1:
                start_date = st.date_input(
                    "Start Date From",
                    value=min_date,
                    min_value=min_date,
                    max_value=max_date,
                    key="start_date_filter"
                )
            with col2:
                end_date = st.date_input(
                    "End Date To",
                    value=max_date,
                    min_value=min_date,
                    max_value=max_date,
                    key="end_date_filter"
                )

            filters['date_range'] = (start_date, end_date)

    # Sorting options
    st.sidebar.markdown("---")
    st.sidebar.subheader("Sorting Options")

    sort_options = ['COURSE CODE', 'CLASS SECTION', 'OFFER DEPT',
                    'ACAD_CAREER', 'START DATE', 'WEEKDAY', 'START TIME', 'END TIME']
    available_sort_options = [opt for opt in sort_options if opt in df.columns]
    sort_by = st.sidebar.selectbox(
        "Sort by", available_sort_options, key="sort_by")
    sort_order = st.sidebar.radio(
        "Order", ['Ascending', 'Descending'], key="sort_order")

    return filters, sort_by, sort_order


def create_current_classes_filters_panel(df):
    """Create filter panel for current classes page"""
    st.sidebar.header("🔍 Current Classes Filters")

    filters = {}

    # Course Code filter
    if 'COURSE CODE' in df.columns:
        course_codes = df['COURSE CODE'].unique()
        course_codes = [code for code in course_codes if code !=
                        '' and pd.notna(code) and code != 'nan']
        course_codes = ['All'] + sorted(course_codes)
        selected_course = st.sidebar.selectbox(
            "Course Code",
            course_codes,
            key="current_course_code_filter"
        )
        if selected_course != 'All':
            filters['COURSE CODE'] = selected_course

    # Class Section filter
    if 'CLASS SECTION' in df.columns:
        class_sections = df['CLASS SECTION'].unique()
        class_sections = [section for section in class_sections if section != '' and pd.notna(
            section) and section != 'nan']
        class_sections = ['All'] + sorted(class_sections)
        selected_section = st.sidebar.selectbox(
            "Class Section",
            class_sections,
            key="current_class_section_filter"
        )
        if selected_section != 'All':
            filters['CLASS SECTION'] = selected_section

    # Curriculum filter
    if 'ACAD_CAREER' in df.columns:
        curricula = df['ACAD_CAREER'].unique()
        curricula = [cur for cur in curricula if cur !=
                     '' and pd.notna(cur) and cur != 'nan']
        curricula = ['All'] + sorted(curricula)
        selected_curriculum = st.sidebar.selectbox(
            "Curriculum",
            curricula,
            key="current_curriculum_filter"
        )
        if selected_curriculum != 'All':
            filters['ACAD_CAREER'] = selected_curriculum

    # Department filter (OFFER DEPT)
    if 'OFFER DEPT' in df.columns:
        departments = df['OFFER DEPT'].unique()
        departments = [dept for dept in departments if dept !=
                       '' and pd.notna(dept) and dept != 'nan']
        departments = ['All'] + sorted(departments)
        selected_department = st.sidebar.selectbox(
            "Department",
            departments,
            key="current_department_filter"
        )
        if selected_department != 'All':
            filters['OFFER DEPT'] = selected_department

    # Venue filter
    if 'VENUE' in df.columns:
        venues = df['VENUE'].unique()
        venues = [v for v in venues if v != '' and pd.notna(v) and v != 'nan']
        venues = ['All'] + sorted(venues)
        selected_venue = st.sidebar.selectbox(
            "Venue",
            venues,
            key="current_venue_filter"
        )
        if selected_venue != 'All':
            filters['VENUE'] = selected_venue

    # Venue is provided checkbox
    st.sidebar.markdown("---")
    venue_provided = st.sidebar.checkbox(
        "Venue is provided", value=True, key="current_venue_provided_filter")
    filters['venue_provided'] = venue_provided

    return filters


def create_upcoming_classes_filters_panel(df):
    """Create filter panel for upcoming classes page"""
    st.sidebar.header("🔍 Upcoming Classes Filters")

    filters = {}

    # Course Code filter
    if 'COURSE CODE' in df.columns:
        course_codes = df['COURSE CODE'].unique()
        course_codes = [code for code in course_codes if code !=
                        '' and pd.notna(code) and code != 'nan']
        course_codes = ['All'] + sorted(course_codes)
        selected_course = st.sidebar.selectbox(
            "Course Code",
            course_codes,
            key="upcoming_course_code_filter"
        )
        if selected_course != 'All':
            filters['COURSE CODE'] = selected_course

    # Class Section filter
    if 'CLASS SECTION' in df.columns:
        class_sections = df['CLASS SECTION'].unique()
        class_sections = [section for section in class_sections if section != '' and pd.notna(
            section) and section != 'nan']
        class_sections = ['All'] + sorted(class_sections)
        selected_section = st.sidebar.selectbox(
            "Class Section",
            class_sections,
            key="upcoming_class_section_filter"
        )
        if selected_section != 'All':
            filters['CLASS SECTION'] = selected_section

    # Curriculum filter
    if 'ACAD_CAREER' in df.columns:
        curricula = df['ACAD_CAREER'].unique()
        curricula = [cur for cur in curricula if cur !=
                     '' and pd.notna(cur) and cur != 'nan']
        curricula = ['All'] + sorted(curricula)
        selected_curriculum = st.sidebar.selectbox(
            "Curriculum",
            curricula,
            key="upcoming_curriculum_filter"
        )
        if selected_curriculum != 'All':
            filters['ACAD_CAREER'] = selected_curriculum

    # Department filter (OFFER DEPT)
    if 'OFFER DEPT' in df.columns:
        departments = df['OFFER DEPT'].unique()
        departments = [dept for dept in departments if dept !=
                       '' and pd.notna(dept) and dept != 'nan']
        departments = ['All'] + sorted(departments)
        selected_department = st.sidebar.selectbox(
            "Department",
            departments,
            key="upcoming_department_filter"
        )
        if selected_department != 'All':
            filters['OFFER DEPT'] = selected_department

    # Venue filter
    if 'VENUE' in df.columns:
        venues = df['VENUE'].unique()
        venues = [v for v in venues if v != '' and pd.notna(v) and v != 'nan']
        venues = ['All'] + sorted(venues)
        selected_venue = st.sidebar.selectbox(
            "Venue",
            venues,
            key="upcoming_venue_filter"
        )
        if selected_venue != 'All':
            filters['VENUE'] = selected_venue

    # Venue is provided checkbox
    st.sidebar.markdown("---")
    venue_provided = st.sidebar.checkbox(
        "Venue is provided", value=True, key="upcoming_venue_provided_filter")
    filters['venue_provided'] = venue_provided

    return filters


def apply_filters(df, filters):
    """Apply filters to dataframe"""
    filtered_df = df.copy()

    # Apply column filters
    for column, value in filters.items():
        if column not in ['date_range', 'time_range', 'venue_provided']:
            filtered_df = filtered_df[filtered_df[column] == value]

    # Apply venue provided filter
    if 'venue_provided' in filters and filters['venue_provided']:
        filtered_df = filtered_df[(filtered_df['VENUE'] != '') & (
            filtered_df['VENUE'].notna()) & (filtered_df['VENUE'] != 'nan')]

    # Apply date range filter
    if 'date_range' in filters:
        start_date, end_date = filters['date_range']
        # Filter out rows with NaT dates before comparison
        date_filtered = filtered_df[filtered_df['START DATE'].notna(
        ) & filtered_df['END DATE'].notna()]
        filtered_df = date_filtered[
            (date_filtered['START DATE'].dt.date >= start_date) &
            (date_filtered['END DATE'].dt.date <= end_date)
        ]

    # Apply time range filter
    if 'time_range' in filters:
        start_time_filter, end_time_filter = filters['time_range']
        # Filter classes that start after or at the selected start time AND end before or at the selected end time
        time_filtered = filtered_df[filtered_df['START TIME'].notna(
        ) & filtered_df['END TIME'].notna()]
        filtered_df = time_filtered[
            (time_filtered['START TIME'] >= start_time_filter) &
            (time_filtered['END TIME'] <= end_time_filter)
        ]

    return filtered_df


def apply_sorting(df, sort_by, sort_order):
    """Apply sorting to dataframe"""
    ascending = sort_order == 'Ascending'

    # For string columns, handle empty strings and NaN
    if sort_by in ['COURSE CODE', 'CLASS SECTION', 'OFFER DEPT', 'ACAD_CAREER', 'WEEKDAY']:
        # Create a temporary column for sorting that handles empty values
        temp_df = df.copy()
        temp_df[f'{sort_by}_sort'] = temp_df[sort_by].replace(
            '', 'zzzz')  # Push empty strings to end
        temp_df = temp_df.sort_values(
            by=f'{sort_by}_sort', ascending=ascending, na_position='last')
        return temp_df.drop(columns=[f'{sort_by}_sort'])
    else:
        # For date/time columns, put NaT values at the end
        return df.sort_values(by=sort_by, ascending=ascending, na_position='last')


def format_time(time_obj):
    """Format time object to string"""
    if pd.isna(time_obj) or time_obj is None or time_obj == '':
        return "N/A"
    if isinstance(time_obj, str):
        return time_obj
    return time_obj.strftime('%H:%M')


def format_date(date_obj):
    """Format date object to string"""
    if pd.isna(date_obj) or date_obj is None or date_obj == '':
        return "N/A"
    if isinstance(date_obj, str):
        return date_obj
    return date_obj.strftime('%Y-%m-%d')


def format_class_number(class_num):
    """Format class number as integer"""
    if pd.isna(class_num) or class_num is None or class_num == '':
        return "N/A"
    try:
        return str(int(float(class_num)))
    except:
        return str(class_num)


def display_course_item(row):
    """Display a single course item in a compact format"""
    with st.container():
        # Create a more compact layout using columns with smaller spacing
        col1, col2, col3 = st.columns([3, 3, 3])
        cc = {
            "NAWD": "green",
            "RPG": "violet",
            "TPG": "red",
            "UG": "blue",
            "UGDE": "blue",
            "UGME": "blue",
        }
        wc = {
            'MON': "red",
            'TUE': "orange",
            'WED': "yellow",
            'THU': "green",
            'FRI': "blue",
            'SAT': "violet",
            'SUN': "gray"
        }
        with col1:
            course_code = row.get('COURSE CODE', 'N/A')
            course_title = row.get('COURSE TITLE', 'N/A')
            st.markdown(f"**{course_code} - {course_title}**")
            class_section = row.get('CLASS SECTION', 'N/A')
            class_number = format_class_number(row.get('CLASS NUMBER', 'N/A'))
            curriculum = row.get('ACAD_CAREER', 'N/A')
            st.markdown(
                f"Section {class_section} • Class No. {class_number} • **:{cc[curriculum]}[{curriculum}]**")

        with col2:
            weekday = row.get('WEEKDAY', 'N/A')
            start_time = format_time(row.get('START TIME'))
            end_time = format_time(row.get('END TIME'))
            venue = row.get('VENUE', 'N/A')
            st.markdown(
                f"**:{wc[weekday]}[{weekday}] {start_time}-{end_time}** @ Venue: **{venue}**")
            start_date = format_date(row.get('START DATE'))
            end_date = format_date(row.get('END DATE'))
            st.markdown(f"**Date:** {start_date} to {end_date}")

        with col3:
            dept = row.get('OFFER DEPT', 'N/A')
            st.markdown(f"{dept}")
            # Truncate instructors if too long
            instructors = row.get('INSTRUCTOR', 'N/A')
            if instructors != 'N/A' and len(str(instructors)) > 25:
                instructors = str(instructors)[:25] + "..."
            st.markdown(f"**Instructor:** {instructors}")

        # Thin separator line
        st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)


def get_current_classes(df):
    """Get classes that are currently running (started but not ended) in HK time"""
    now_hk = get_hk_time()
    current_time = now_hk.time()
    current_date = now_hk.date()

    # Map weekday names
    weekday_map = {
        'Monday': 'MON',
        'Tuesday': 'TUE',
        'Wednesday': 'WED',
        'Thursday': 'THU',
        'Friday': 'FRI',
        'Saturday': 'SAT',
        'Sunday': 'SUN'
    }
    current_weekday = weekday_map[now_hk.strftime('%A')]

    # Filter for today's classes on current weekday
    today_classes = df[
        (df['WEEKDAY'] == current_weekday) &
        (df['START DATE'].dt.date <= current_date) &
        (df['END DATE'].dt.date >= current_date)
    ].copy()

    # Filter for classes that are currently running
    current_classes = today_classes[
        (today_classes['START TIME'] <= current_time) &
        (today_classes['END TIME'] >= current_time)
    ]

    return current_classes


def get_upcoming_classes_today(df):
    """Get classes that are upcoming today (haven't started yet) in HK time"""
    now_hk = get_hk_time()
    current_time = now_hk.time()
    current_date = now_hk.date()

    # Map weekday names
    weekday_map = {
        'Monday': 'MON',
        'Tuesday': 'TUE',
        'Wednesday': 'WED',
        'Thursday': 'THU',
        'Friday': 'FRI',
        'Saturday': 'SAT',
        'Sunday': 'SUN'
    }
    current_weekday = weekday_map[now_hk.strftime('%A')]

    # Filter for today's classes on current weekday
    today_classes = df[
        (df['WEEKDAY'] == current_weekday) &
        (df['START DATE'].dt.date <= current_date) &
        (df['END DATE'].dt.date >= current_date)
    ].copy()

    # Filter for classes that haven't started yet today
    upcoming_classes = today_classes[today_classes['START TIME']
                                     > current_time]

    return upcoming_classes


def display_current_classes_page(df):
    """Display page for currently running classes"""
    st.header("🏫 Currently Running Classes")

    # Get current date and time info in HK timezone
    now_hk = get_hk_time()
    current_time = now_hk.strftime('%H:%M')
    current_date = now_hk.strftime('%Y-%m-%d')
    current_weekday = now_hk.strftime('%A')
    timezone_info = "GMT+8"

    st.markdown(
        f"**Current Time (Hong Kong {timezone_info}):** {current_time} | **Date:** {current_date} | **Day:** {current_weekday}")

    # Refresh button
    col1, col2, col3 = st.columns([3, 3, 3])
    with col1:
        if st.button("🔄 Refresh Current Time", type="primary", use_container_width=True):
            st.rerun()

    st.markdown("---")

    # Get current classes
    current_classes = get_current_classes(df)

    # Create filters for current classes
    filters = create_current_classes_filters_panel(df)

    # Apply filters
    if filters:
        current_classes = apply_filters(current_classes, filters)

    # Sorting options for current classes
    st.sidebar.markdown("---")
    st.sidebar.subheader("Current Classes Sorting")

    sort_options = ['COURSE CODE', 'CLASS SECTION',
                    'OFFER DEPT', 'ACAD_CAREER', 'START TIME', 'END TIME']
    available_sort_options = [opt for opt in sort_options if opt in df.columns]
    sort_by = st.sidebar.selectbox(
        "Sort by", available_sort_options, key="current_sort_by")
    sort_order = st.sidebar.radio(
        "Order", ['Ascending', 'Descending'], key="current_sort_order")

    # Apply sorting
    current_classes = apply_sorting(current_classes, sort_by, sort_order)

    # Display results
    if current_classes.empty:
        st.info("No classes are currently running.")
        st.markdown("💡 *All classes have either ended or haven't started yet.*")
    else:
        st.markdown(
            f"**:green[Found {len(current_classes)} classes currently running]**")

        # Pagination
        if 'current_display_count' not in st.session_state:
            st.session_state.current_display_count = 20

        display_df = current_classes.head(
            st.session_state.current_display_count)

        for idx, row in display_df.iterrows():
            display_course_item(row)

        # Load more button
        if len(current_classes) > st.session_state.current_display_count:
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button(f"Load More ({min(20, len(current_classes) - st.session_state.current_display_count)} more available)", key="current_load_more"):
                    st.session_state.current_display_count += 20
                    st.rerun()


def display_upcoming_classes_page(df):
    """Display page for upcoming classes today"""
    st.header("⏰ Upcoming Classes Today")

    # Get current date and time info in HK timezone
    now_hk = get_hk_time()
    current_time = now_hk.strftime('%H:%M')
    current_date = now_hk.strftime('%Y-%m-%d')
    current_weekday = now_hk.strftime('%A')
    timezone_info = "GMT+8"

    st.markdown(
        f"**Current Time (Hong Kong {timezone_info}):** {current_time} | **Date:** {current_date} | **Day:** {current_weekday}")

    # Refresh button
    col1, col2, col3 = st.columns([3, 3, 3])
    with col1:
        if st.button("🔄 Refresh Current Time", type="primary", use_container_width=True):
            st.rerun()

    st.markdown("---")

    # Get upcoming classes
    upcoming_classes = get_upcoming_classes_today(df)

    # Create filters for upcoming classes
    filters = create_upcoming_classes_filters_panel(df)

    # Apply filters
    if filters:
        upcoming_classes = apply_filters(upcoming_classes, filters)

    # Sorting options for upcoming classes
    st.sidebar.markdown("---")
    st.sidebar.subheader("Upcoming Classes Sorting")

    sort_options = ['START TIME', 'END TIME', 'COURSE CODE',
                    'CLASS SECTION', 'OFFER DEPT', 'ACAD_CAREER']
    available_sort_options = [opt for opt in sort_options if opt in df.columns]
    sort_by = st.sidebar.selectbox(
        "Sort by", available_sort_options, key="upcoming_sort_by")
    sort_order = st.sidebar.radio(
        "Order", ['Ascending', 'Descending'], key="upcoming_sort_order")

    # Apply sorting
    upcoming_classes = apply_sorting(upcoming_classes, sort_by, sort_order)

    # Display results
    if upcoming_classes.empty:
        st.info("No more classes scheduled for today.")
        st.markdown("🎉 *You're done for the day!*")
    else:
        st.markdown(
            f"**:green[Found {len(upcoming_classes)} upcoming classes today]**")

        # Pagination
        if 'upcoming_display_count' not in st.session_state:
            st.session_state.upcoming_display_count = 20

        display_df = upcoming_classes.head(
            st.session_state.upcoming_display_count)

        for idx, row in display_df.iterrows():
            display_course_item(row)

        # Load more button
        if len(upcoming_classes) > st.session_state.upcoming_display_count:
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button(f"Load More ({min(20, len(upcoming_classes) - st.session_state.upcoming_display_count)} more available)", key="upcoming_load_more"):
                    st.session_state.upcoming_display_count += 20
                    st.rerun()


def main():
    # Load data first to get the term
    df = load_and_preprocess_data()

    if df.empty:
        st.warning(
            "No data loaded. Please check if the Excel file exists in the working directory.")
        return

    # Display term information at the top
    if 'TERM' in df.columns and not df.empty:
        term = df['TERM'].iloc[0] if not df['TERM'].isna(
        ).all() else "Unknown Term"
        st.title(f"📚 HKU Course Explorer - {term}")
    else:
        st.title("📚 HKU Course Explorer")

    # Navigation
    st.sidebar.markdown("---")
    st.sidebar.subheader("Navigation")
    page = st.sidebar.radio(
        "Go to:",
        ["All Courses", "Current Classes", "Upcoming Classes Today"],
        key="navigation"
    )

    if page == "All Courses":
        st.markdown(
            "Browse and filter course schedules from the uploaded Excel file.")

        # Display data info
        st.sidebar.markdown(f"**Total records:** {len(df)}")

        # Create filters panel
        filters, sort_by, sort_order = create_filters_panel(df)

        # Apply filters and sorting
        filtered_df = apply_filters(df, filters)
        filtered_df = apply_sorting(filtered_df, sort_by, sort_order)

        # Display results count
        st.markdown(f"**:green[Found {len(filtered_df)} matching records]**")

        # Pagination
        if 'display_count' not in st.session_state:
            st.session_state.display_count = 20

        # Display courses
        display_df = filtered_df.head(st.session_state.display_count)

        if display_df.empty:
            st.info("No courses match your filters. Try adjusting your criteria.")
        else:
            for idx, row in display_df.iterrows():
                display_course_item(row)

        # Load more button
        if len(filtered_df) > st.session_state.display_count:
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button(f"Load More ({min(20, len(filtered_df) - st.session_state.display_count)} more available)"):
                    st.session_state.display_count += 20
                    st.rerun()

        # Reset display count when filters change
        if st.session_state.get('previous_filter_count', 0) != len(filtered_df):
            st.session_state.display_count = 20
            st.session_state.previous_filter_count = len(filtered_df)

    elif page == "Current Classes":
        display_current_classes_page(df)

    elif page == "Upcoming Classes Today":
        display_upcoming_classes_page(df)


if __name__ == "__main__":
    main()
