            # Helper functions to convert seconds and hours to HH:MM:SS format
            def seconds_to_hms(seconds):
                if pd.isna(seconds) or seconds == 0:
                    return "00:00:00"
                hours = int(seconds // 3600)
                minutes = int((seconds % 3600) // 60)
                seconds = int(seconds % 60)
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            
            # Helper function to convert decimal hours to HH:MM:SS format
            def hours_to_hms(decimal_hours):
                if pd.isna(decimal_hours) or decimal_hours == 0:
                    return "00:00:00"
                total_seconds = decimal_hours * 3600
                hours = int(total_seconds // 3600)
                minutes = int((total_seconds % 3600) // 60)
                seconds = int(total_seconds % 60)
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
