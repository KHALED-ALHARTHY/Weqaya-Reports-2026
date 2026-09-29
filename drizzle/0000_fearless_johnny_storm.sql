CREATE TABLE `entries` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`year` integer NOT NULL,
	`program` text NOT NULL,
	`facility_id` text NOT NULL,
	`period` text NOT NULL,
	`score` real NOT NULL,
	`numerator` integer,
	`denominator` integer,
	`observation` text DEFAULT '' NOT NULL,
	`action` text DEFAULT '' NOT NULL,
	`owner` text DEFAULT '' NOT NULL,
	`due_date` text DEFAULT '' NOT NULL,
	`updated_at` text NOT NULL
);
--> statement-breakpoint
CREATE UNIQUE INDEX `entries_year_program_facility_period_unique` ON `entries` (`year`,`program`,`facility_id`,`period`);--> statement-breakpoint
CREATE INDEX `entries_year_program_idx` ON `entries` (`year`,`program`);