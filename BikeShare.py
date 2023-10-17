import time
import pandas as pd
import numpy as np

CITY_DATA = { 'ch': 'chicago.csv',
              'ny': 'new_york_city.csv',
              'ws': 'washington.csv' }



def get_filters():

        # ask user to choose city
    city = input("Kindly select a city to display statistics about.\n \
        You can type (ny) for new york city,(ws) for washington,(ch) for chicago. \n").lower()

         # Validating city input
    while city not in CITY_DATA :
            print("that's invalid input")

            city = input("Kindly select a city to display statistics about.\n \
        You can type (ny) for new york city,(ws) for washington,(ch) for chicago. \n").lower()



    print('Hello! Let\'s explore some US bikeshare data!')





            # ask user to choose Month
    month = input("Kindly select a month to display statistics about.\n \
        choose one from this list (january),(february),(march),(april),(may),(june) \n").lower()

         # Validating month input
    while month not in ['january', 'february', 'march', 'april', 'may', 'june', 'all']:
            print("that's invalid input, ")

            month = input("Kindly select a month to display statistics about.\n \
        choose one from this list (january),(february),(march),(april),(may),(june) \n").lower()





                # ask user to choose day
    day = input("Kindly select a day to display statistics about.\n \
        choose one from this list (Sunday),(Monday),(Tuesday),(Wednesday),(Thursday),(Friday),(Saturday),(all) \n").lower()

         # Validating day input
    while day not in ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'all']:
            print("that's invalid input, ")

            day = input("Kindly select a day to display statistics about.\n \
        choose one from this list (Sunday),(Monday),(Tuesday),(Wednesday),(Thursday),(Friday),(Saturday),(all) \n").lower()



    print('-'*40)
    return city, month, day


def load_data(city, month, day):
       # load data file into a dataframe
    df = pd.read_csv(CITY_DATA[city])

    # convert the Start Time column to datetime
    df['Start Time'] = pd.to_datetime(df['Start Time'])

    # extract month and day of week from Start Time to create new columns
    df['month'] = df['Start Time'].dt.month
    df['day_of_week'] = df['Start Time'].dt.day_name

    # filter by month if applicable
    if month != 'all':
        # use the index of the months list to get the corresponding int
        months = ['january', 'february', 'march', 'april', 'may', 'june']
        month = months.index(month) + 1

        # filter by month to create the new dataframe
        df = df[df['month'] == month]

    # filter by day of week if applicable
    if day != 'all':
        # filter by day of week to create the new dataframe
        df = df[df['day_of_week'] == day.title()]

    return df


def time_stats(df):
    """Displays statistics on the most frequent times of travel."""

    print('\nCalculating The Most Frequent Times of Travel...\n')
    start_time = time.time()

   # display the most common month
    df['month'] = df['Start Time'].dt.month
    common_month = df['month'].mode()[0]
    print('most common month is : ', common_month)
    # display the most common day of week
    df['day_of_week'] = df['Start Time'].dt.week
    common_day = df['day_of_week'].value_counts()
    print('most common day is : ', common_day)

    # display the most common start hour
    df['hour'] = df['Start Time'].dt.hour
    common_hour = df['hour'].mode()[0]
    print('most common hour is : ', common_hour)

    print("\nThis took %s seconds." % (time.time() - start_time))
    print('-'*40)


def station_stats(df):
    """Displays statistics on the most popular stations and trip."""

    print('\nCalculating The Most Popular Stations and Trip...\n')
    start_time = time.time()

    # display most commonly used start station
    common_start = df['Start Station'].mode()[0]
    print('most common start station is : ', common_start)
    # display most commonly used end station
    common_end = df['End Station'].mode()[0]
    print('most common end station is : ', common_end)
    # display most frequent combination of start station and end station trip
    df['combination'] = df['Start Station'] + '--' + df['End Station']
    common_route = df['combination'].mode()[0]
    print('most common route is : ', common_route)
    print("\nThis took %s seconds." % (time.time() - start_time))
    print('-'*40)


def trip_duration_stats(df):
    """Displays statistics on the total and average trip duration."""


    print('\nCalculating Trip Duration...\n')
    start_time = time.time()

    # display total travel time
    total_travel = df['Trip Duration'].sum()
    Myday = total_travel // (24*3600)
    total_travel = total_travel % (24*3600)
    MyHour = total_travel //3600
    total_travel %=  3600
    MyMinuits = total_travel //60
    total_travel %= 60
    MySeconds = total_travel

    print('total travel time is : {} Days, {} Hours {} minuits {} Seconds'.format(Myday, MyHour, MyMinuits, MySeconds))

      # display mean travel time
    mean_travel = df['Trip Duration'].mean()

    Myday = mean_travel // (24*3600)
    mean_travel = mean_travel % (24*3600)
    MyHour = mean_travel //3600
    mean_travel %=  3600
    MyMinuits = mean_travel //60
    mean_travel %= 60
    MySeconds = mean_travel

    print('mean travel time is : {} Days, {} Hours {} minuits {} Seconds'.format(Myday, MyHour, MyMinuits, MySeconds))
    print("\nThis took %s seconds." % (time.time() - start_time))
    print('-'*40)


def user_stats(df):
    """Displays statistics on bikeshare users."""

    print('\nCalculating User Stats...\n')
    start_time = time.time()

    # Display counts of user types
    user_types = df['User Type'].value_counts()
    print(user_types)

    # Display counts of gender
    if 'Gender' in df:
        gender = df['Gender'].value_counts()
        print(gender)
    else:
        print("There is no gender information in this city.")

    # Display earliest, most recent, and most common year of birth
    if 'Birth_Year' in df:
        earliest = df['Birth_Year'].min()
        print(earliest)
        recent = df['Birth_Year'].max()
        print(recent)
        common_birth = df['Birth Year'].mode()[0]
        print(common_birth)
    else:
        print("There is no birth year information in this city.")

    print("\nThis took %s seconds." % (time.time() - start_time))
    print('-'*40)

def show_data_rows(df):

    i = 0
    while True :
        print(df.iloc[i:i+5])
        answer = input('Show 5 Rows of Data ?? (y) or (n)').lower()
        i += 5
        if answer !="y" :
            break


def main():
    while True:
        city, month, day = get_filters()
        df = load_data(city, month, day)

        time_stats(df)
        station_stats(df)
        trip_duration_stats(df)
        user_stats(df)
        show_data_rows(df)

        restart = input('\nStart Again? Enter (y) or (n).\n')
        if restart.lower() != 'y':
            break


if __name__ == "__main__":
    main()



    
