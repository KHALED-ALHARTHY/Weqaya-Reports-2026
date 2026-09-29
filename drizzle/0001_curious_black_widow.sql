CREATE TABLE `report_settings` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`year` integer NOT NULL,
	`scope` text NOT NULL,
	`program_filter` text NOT NULL,
	`title` text NOT NULL,
	`organization` text NOT NULL,
	`executive_summary` text NOT NULL,
	`methodology` text NOT NULL,
	`conclusion` text NOT NULL,
	`prepared_by` text DEFAULT '' NOT NULL,
	`reviewed_by` text DEFAULT '' NOT NULL,
	`approved_by` text DEFAULT '' NOT NULL,
	`program_notes_json` text DEFAULT '{}' NOT NULL,
	`include_methodology` integer DEFAULT true NOT NULL,
	`include_facility_details` integer DEFAULT true NOT NULL,
	`include_notes` integer DEFAULT true NOT NULL,
	`updated_at` text NOT NULL
);
--> statement-breakpoint
CREATE UNIQUE INDEX `report_settings_year_scope_program_unique` ON `report_settings` (`year`,`scope`,`program_filter`);