from weekendWorkTracker import RunProgram

def main():
	
	initiate = RunProgram()

	initiate.config_setup()
	initiate.create_working_list()
	initiate.populate_working_list()
	initiate.get_current_date()
	initiate.get_next_weekend()
	initiate.format_wkst_title_dates()
	initiate.setup_excel()
	initiate.save_workbook()
	initiate.open_excel()
	

if __name__ == "__main__":
	main()