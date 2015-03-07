require 'yaml'
require 'set'

# options to implement: (1) save current registry keys, (2) backup peak shift values to a .reg file
# (3) check teco version on exe, check registry key presence, (4) decide where to save .reg file, dump log information on screen

class HolidayManager
	def initialize(weekend, additional_holidays, time)
		@weekend = Set.new weekend
		@additional_holidays = Set.new additional_holidays
		@time = time
	end

	def getNextWorkBegin
		while @weekend.include? @time.wday or @additional_holidays.include? @time
			@time = @time + 1
		end

		return @time
	end

	def getNextWorkEnd
		while ! @weekend.include? @time.wday and ! @additional_holidays.include? @time
			@time = @time + 1
		end
		return @time - 1
	end
end

def printPreamble()
	puts <<-EOF
	Windows Registry Editor Version 5.00

	[HKEY_LOCAL_MACHINE\\SOFTWARE\\Toshiba\\eco Utility\\PeakShift]
	"NotifyTime"=dword:0000003c
	"IsSupport"=dword:00000001
	"PeakShiftCount"=dword:00000004
	"IsNotify"=dword:00000001
	EOF
end

yamlProps = YAML.load_file("teco_peakshift_settings.yaml")
additional_holidays = yamlProps["additional_holidays"]
weekend = yamlProps["weekend"]

printPreamble()

startHour   = 8
startMinute = 0
endHour     = 19
endMinute   = 5
minCharge   = 15
active      = 1

holidayManager = HolidayManager.new(weekend, additional_holidays, Date.today)

for i in 1..4 do
	startTime = holidayManager.getNextWorkBegin
	endTime = holidayManager.getNextWorkEnd

	puts "startTime: #{startTime}, endTime: #{endTime}"

	startMonth  = startTime.month
	startDay    = startTime.day
	endMonth    = endTime.month
	endDay      = endTime.day

	puts <<-EOF
	"PeakShift#{i}"=hex: \\
					  #{startMonth.to_s(16)},00,00,00, \\
					    #{startDay.to_s(16)},00,00,00, \\
					    #{endMonth.to_s(16)},00,00,00, \\
					      #{endDay.to_s(16)},00,00,00, \\
					   #{startHour.to_s(16)},00,00,00, \\
					 #{startMinute.to_s(16)},00,00,00, \\
					     #{endHour.to_s(16)},00,00,00, \\
					   #{endMinute.to_s(16)},00,00,00, \\
					   #{minCharge.to_s(16)},00,00,00, \\
					      #{active.to_s(16)},00,00,00
	EOF
end