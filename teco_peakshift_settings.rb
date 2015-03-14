require 'yaml'
require 'set'
require 'logger'

# options to implement:
# (*) backup peak shift values to a .reg file
# => reg export "HKEY_LOCAL_MACHINE\\SOFTWARE\\Toshiba\\eco Utility\\PeakShift" backup.reg
# (*) check teco version on exe, check registry key presence
# (*) decide where to save .reg file, dump log information on screen

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

def printPreamble(registryFile)
	registryFile.puts <<-EOF
	Windows Registry Editor Version 5.00

	[HKEY_LOCAL_MACHINE\\SOFTWARE\\Toshiba\\eco Utility\\PeakShift]
	"NotifyTime" = dword:0000003c
	"IsSupport" = dword:00000001
	"PeakShiftCount" = dword:00000004
	"IsNotify" = dword:00000001
	EOF
end

ARGV[0] or abort "Usage: #{__FILE__} <file.reg>\nSettings must be in file teco_peakshift_settings.yaml"

YamlProps = YAML.load_file("teco_peakshift_settings.yaml")
Additional_holidays = YamlProps["additional_holidays"]
Weekend     = YamlProps["weekend"]
StartHour   = YamlProps["startHour"]
StartMinute = YamlProps["startMinute"]
EndHour     = YamlProps["endHour"]
EndMinute   = YamlProps["endMinute"]
MinCharge   = YamlProps["minCharge"]

Active      = 1

holidayManager = HolidayManager.new(Weekend, Additional_holidays, Date.today)

@log = Logger.new(STDOUT)
@log.level = Logger::INFO

if (@log.level != Logger::DEBUG)
	@log.formatter = proc do |severity, datetime, progname, msg|
		"#{msg}\n"
	end
end

File.open("#{ARGV[0]}", "w") do |registryFile|
	@log.info "Creating registry file #{ARGV[0]}"

	printPreamble(registryFile)
	for i in 1..4 do
		startTime = holidayManager.getNextWorkBegin
		endTime = holidayManager.getNextWorkEnd

		@log.info "Week number #{i}, start: #{startTime}, end: #{endTime}"

		startMonth  = startTime.month
		startDay    = startTime.day
		endMonth    = endTime.month
		endDay      = endTime.day

		registryFile.puts <<-EOF
		"PeakShift#{i}"=hex: \\
						  #{startMonth.to_s(16)},00,00,00, \\
							#{startDay.to_s(16)},00,00,00, \\
							#{endMonth.to_s(16)},00,00,00, \\
							  #{endDay.to_s(16)},00,00,00, \\
						   #{StartHour.to_s(16)},00,00,00, \\
						 #{StartMinute.to_s(16)},00,00,00, \\
							 #{EndHour.to_s(16)},00,00,00, \\
						   #{EndMinute.to_s(16)},00,00,00, \\
						   #{MinCharge.to_s(16)},00,00,00, \\
							  #{Active.to_s(16)},00,00,00
		EOF
	end
	@log.info "Registry file #{ARGV[0]} created, now please import it with regedit.exe"
end
