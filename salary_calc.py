from datetime import datetime, timedelta

class SalaryCalculator:
    def __init__(self):  # Corrected constructor
        self.daily_records = []

    def add_work_day(self, date, start_time, end_time, in_control_room,
                     is_friday, is_saturday, is_holiday, is_last_day_of_holiday):
        base_rate = 60 if in_control_room else 50

        # Combine date and time
        start_datetime = datetime.combine(date, start_time)
        end_datetime = datetime.combine(date, end_time)

        # Handle night shifts that cross midnight
        if end_datetime <= start_datetime:
            end_datetime += timedelta(days=1)

        # Edge case: if start_time is equal to end_time, raise an error
        if start_datetime == end_datetime:
            raise ValueError("Start time cannot be equal to end time")

        total_pay = 0.0

        # Calculate pay for each hour
        current_time = start_datetime
        while current_time < end_datetime:
            next_hour = current_time + timedelta(hours=1)
            if next_hour > end_datetime:
                next_hour = end_datetime  # Adjust for partial hours

            hour_duration = (next_hour - current_time).total_seconds() / 3600
            multiplier = self._get_multiplier(current_time, is_friday, is_saturday,
                                              is_holiday, is_last_day_of_holiday)
            total_pay += base_rate * multiplier * hour_duration

            current_time = next_hour  # Move the loop forward

        # Append the record to daily_records with all relevant information
        self.daily_records.append({
            'date': date,
            'start_time': start_time,
            'end_time': end_time,
            'pay': total_pay
        })

    def _get_multiplier(self, current_time, is_friday, is_saturday, is_holiday, is_last_day_of_holiday):
        multiplier = 1.0

        # Apply 1.5x multiplier for special conditions
        if is_friday or is_holiday:
            multiplier = 1.5
        elif is_saturday or is_last_day_of_holiday:
            # Apply 1.5x multiplier until 4 AM the next day
            if current_time.hour < 4 or (current_time.hour == 4 and current_time.minute == 0):
                multiplier = 1.5
            else:
                multiplier = 1.0
        else:
            multiplier = 1.0
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
