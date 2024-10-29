from datetime import datetime, timedelta

class SalaryCalculator:
    def __init__(self):
        self.daily_records = []




    def add_work_day(self, date, start_time, end_time, in_control_room, is_holiday_eve,
                     is_friday, is_saturday, is_holiday, is_last_day_of_holiday):
        base_rate = 60 if in_control_room else 50

        # combine date and time
        start_datetime = datetime.combine(date, start_time)
        end_datetime = datetime.combine(date, end_time)

        # night shifts that cross midnight, add a day to the end_datetime.
        if end_datetime <= start_datetime:
            end_datetime += timedelta(days=1)

        # edge case: if start_time is equal to end_time
        if start_datetime == end_datetime:
            raise ValueError("Start time cannot be equal to end time")

        total_pay = 0.0

        # calculate pay for each hour segment
        current_time = start_datetime
        while current_time < end_datetime:
            next_hour = current_time + timedelta(hours=1)
            if next_hour > end_datetime:
                next_hour = end_datetime  # for partial hours

            hour_duration = (next_hour - current_time).total_seconds() / 3600
            multiplier = self._get_correct_multiplier(current_time, is_holiday, is_holiday_eve, is_last_day_of_holiday)
            total_pay += base_rate * multiplier * hour_duration

            current_time = next_hour  # Move the loop forward

        # add the record to daily_records with all relevant information
        self.daily_records.append({
            'date': date,
            'start_time': start_time,
            'end_time': end_time,
            'pay': total_pay
        })




    def _get_correct_multiplier(self, current_time, is_holiday=False, is_holiday_eve=False, is_last_day_of_holiday=False):
        """
        Determine the pay multiplier based on the day and time.
        - If it's Friday, the multiplier is 1.5x for the entire day.
        - If it's Saturday, the multiplier is 1.5x for the entire day.
        - If it's Sunday before 4 AM, the multiplier is 1.5x.
        - For other days and times, the multiplier is 1.0x.
        """
        if is_holiday:
            return 1.5
        
        multiplier = 1.0

        weekday = current_time.weekday()  # Get the day of the week (0 = Monday, 1 = Tuesday, 2 = Wednesday, 3 = Thursday, 4 = Friday, 5 = Saturday, 6 = Sunday)
        
        if (weekday == 4 or is_holiday_eve) and current_time.hour >= 16:
            multiplier = 1.5
        elif weekday == 5:  
            multiplier = 1.5
        elif is_last_day_of_holiday and current_time.hour < 4:
            multiplier = 1.5
        elif weekday == 6 and current_time.hour < 4:  # Sunday before 4am
            multiplier = 1.5

        return multiplier





    def total_pay(self):
        return sum(record['pay'] for record in self.daily_records)






    def show_all_days(self):
        if not self.daily_records:
            print("No work days have been added.")
            return

        for record in self.daily_records:
            date_str = record['date'].strftime("%Y-%m-%d")
            start_time_str = record['start_time'].strftime("%H:%M")
            end_time_str = record['end_time'].strftime("%H:%M")
            pay = record['pay']
            print(f"Date: {date_str}, Start Time: {start_time_str}, End Time: {end_time_str}, Pay: {pay:.2f} shekels")






    def calculate_travel_charge(self, date, start_time, end_time, is_friday, is_saturday):
        """
        Calculate the travel charge based on the day and time.
        """
        travel_charge = 12  
        
        # friday charge 
        if is_friday:
            if start_time < datetime.strptime("15:00", "%H:%M").time():
                travel_charge = 12  # 6 shekels each way
            else:
                travel_charge = 40  # 40 shekels after 3:00 p.m for both way, 20 each.
        
        # saturday charge
        elif is_saturday:
            if start_time >= datetime.strptime("23:00", "%H:%M").time() and end_time <= datetime.strptime("07:00", "%H:%M").time():
                travel_charge = 20 + 6  # 20 shekels to work, 6 shekels return.
            else:
                travel_charge = 12  # 12 shekels total.
        
        return travel_charge