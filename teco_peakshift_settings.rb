require 'yaml'
require 'set'
require 'logger'
require 'getoptlong'

@registryAccess = true
begin
	require 'win32/registry'
rescue LoadError
	@registryAccess = false
end

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

def checkRegistryKeyPresence
	@registryAccess or return
	begin
		# fails if the registry keys that should be present because of "Toshiba Eco Utility" are not there
		Win32::Registry::HKEY_LOCAL_MACHINE.open "SOFTWARE2\\Toshiba\\eco Utility\\PeakShift"
	rescue Win32::Registry::Error
		raise "Peakshift registry key not present"
	end
end

def checkToshibaEcoUtilityPresence
	@registryAccess or return

	begin
		key = Win32::Registry::HKEY_LOCAL_MACHINE.open "SOFTWARE\\Toshiba\\eco Utility"
		installDir = key['InstallDir']
		File.exists? "#{installDir}/Teco.exe" or raise 'it seems like toshiba eco utility is not installed'
	rescue Win32::Registry::Error
		raise "Toshiba eco utility install dir not set"
	end
end

def winRegBackupAndImport
	# TODO: if using windows
	require 'win32ole'
	shell = WIN32OLE.new('Shell.Application')

	# backup
	shell.ShellExecute 'reg', 'export /y "HKEY_LOCAL_MACHINE\\SOFTWARE\\Toshiba\\eco Utility\\PeakShift" teco_backup.reg'

	# import
	# TODO: this works fine with native windows ruby versions, not with the cygwin one. Check this
	# shell.ShellExecute 'regedit.exe', "#{Dir.pwd}/#{ARGV[0]}"
	# command forked, can't wait for termination
end

HELP_MSG = <<-EOF
Usage: #{__FILE__} [options] <file.reg>
-h, --help:
	show help

--continue, -c:
	tries to continue even if checking for toshiba eco utility presence was unsuccessful

Settings must be in file teco_peakshift_settings.yaml
EOF
continueAnyway = false

opts = GetoptLong.new(
	[ '--help', '-h', GetoptLong::NO_ARGUMENT ],
	[ '--continue', '-c', GetoptLong::NO_ARGUMENT ]
)
opts.each do |opt,arg|
	case opt
		when '--help'
			puts HELP_MSG
			exit
		when '--continue'
			continueAnyway = true
	end
end

(outputReg = ARGV.shift) or abort HELP_MSG

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

begin
	checkRegistryKeyPresence
	checkToshibaEcoUtilityPresence
rescue => e
	@log.error e
	continueAnyway or puts "execution will stop, please specify command line option '-c' if you really want to continue"
	continueAnyway or raise
end

File.open(outputReg, "w") do |registryFile|
	@log.info "Creating registry file #{outputReg}"

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
	@log.info "Registry file #{outputReg} created, now please import it with regedit.exe"

	winRegBackupAndImport
end
