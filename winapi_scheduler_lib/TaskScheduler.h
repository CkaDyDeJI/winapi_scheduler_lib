#pragma once

class TaskScheduler
{
	bool scheduleOnTime();
	bool scheduleDaily();
	bool scheduleOnBoot();
	bool scheduleWeekly();
	bool scheduleOnReg();
	bool scheduleOnLogon();
	bool scheduleOnIdle();
	bool scheduleCustom();
};

