using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace VFormat.Addin;

internal static class RibbonIconRenderer
{
	private enum IconTheme
	{
		HeroCoral,      // Large 3D Glossy Coral-Orange -> Deep Orange badge with white FA icon
		HeroAmber,      // Large 3D Glossy Amber-Gold -> Coral-Orange badge with white FA icon
		HeroTerracotta, // Large 3D Glossy Terracotta -> Coral badge with white FA icon
		CoralDuotone,   // Soft Peach card (#FFF1E6) + Vibrant Coral-Orange (#E04F26) FA icon
		AmberDuotone,   // Soft Amber card (#FEF3C7) + Warm Amber-Orange (#D97706) FA icon
		CrimsonDuotone, // Soft Rose card (#FEE2E2) + Crimson-Coral (#DC2626) FA icon
		EmeraldDuotone  // Soft Mint card (#DCFCE7) + Emerald (#15803D) FA icon
	}

	private sealed class IconSpec
	{
		public string PathData;
		public IconTheme Theme;

		public IconSpec(string pathData, IconTheme theme)
		{
			PathData = pathData;
			Theme = theme;
		}
	}

	// Font Awesome 6 Solid SVG Path Definitions
	private const string FA_WandMagicSparkles = "M24 0C37.3 0 48 10.7 48 24V40c0 13.3 10.7 24 24 24H88c13.3 0 24 10.7 24 24s-10.7 24-24 24H72c-13.3 0-24 10.7-24 24v16c0 13.3-10.7 24-24 24s-24-10.7-24-24V136C0 122.7-10.7 112-24 112H-40c-13.3 0-24-10.7-24-24s10.7-24 24-24h16C-10.7 64 0 53.3 0 40V24C0 10.7 10.7 0 24 0zM439 7c9.4-9.4 24.6-9.4 33.9 0l32 32c9.4 9.4 9.4 24.6 0 33.9l-368 368c-9.4 9.4-24.6 9.4-33.9 0l-32-32c-9.4-9.4-9.4-24.6 0-33.9L439 7zM328 152l32 32L456 88 424 56 328 152zM448 352c13.3 0 24 10.7 24 24v16c0 13.3 10.7 24 24 24h16c13.3 0 24 10.7 24 24s-10.7 24-24 24H496c-13.3 0-24 10.7-24 24v16c0 13.3-10.7 24-24 24s-24-10.7-24-24V488c0-13.3-10.7-24-24-24H384c-13.3 0-24-10.7-24-24s10.7-24 24-24h16c13.3 0 24-10.7 24-24V376c0-13.3 10.7-24 24-24z";
	private const string FA_SpellCheck = "M160 32c-13.1 0-24.6 8.5-28.5 21L53.5 301H24c-13.3 0-24 10.7-24 24s10.7 24 24 24H96c13.3 0 24-10.7 24-24s-10.7-24-24-24h-.8l16-51H208.8l16 51H224c-13.3 0-24 10.7-24 24s10.7 24 24 24h72c13.3 0 24-10.7 24-24s-10.7-24-24-24H266.5L188.5 53C184.6 40.5 173.1 32 160 32zm0 92.2L193.8 202H126.2L160 124.2zM569 271c9.4-9.4 9.4-24.6 0-33.9s-24.6-9.4-33.9 0L376 396.1l-63-63c-9.4-9.4-24.6-9.4-33.9 0s-9.4 24.6 0 33.9l80 80c9.4 9.4 24.6 9.4 33.9 0L569 271z";
	private const string FA_MagnifyingGlass = "M416 208c0 45.9-14.9 88.3-40 122.7L502.6 457.4c12.5 12.5 12.5 32.8 0 45.3s-32.8 12.5-45.3 0L330.7 376c-34.4 25.2-76.8 40-122.7 40C93.1 416 0 322.9 0 208S93.1 0 208 0S416 93.1 416 208zM208 352a144 144 0 1 0 0-288 144 144 0 1 0 0 288z";
	private const string FA_ClipboardCheck = "M192 0c-41.8 0-77.4 26.7-90.5 64H64C28.7 64 0 92.7 0 128V448c0 35.3 28.7 64 64 64H320c35.3 0 64-28.7 64-64V128c0-35.3-28.7-64-64-64H282.5C269.4 26.7 233.8 0 192 0zm0 64a32 32 0 1 1 0 64 32 32 0 1 1 0-64zM305 273L177 401c-9.4 9.4-24.6 9.4-33.9 0L79 337c-9.4-9.4-9.4-24.6 0-33.9s24.6-9.4 33.9 0l47 47L271 239c9.4-9.4 24.6-9.4 33.9 0s9.4 24.6 0 33.9z";
	private const string FA_CommentSlash = "M38.8 5.1C28.4-3.1 13.3-1.2 5.1 9.2S-1.2 34.7 9.2 42.9l592 464c10.4 8.2 25.5 6.3 33.7-4.1s6.3-25.5-4.1-33.7L512.9 376.7C552.2 336.2 576 285.3 576 224C576 100.3 461.4 0 320 0C226.7 0 145 43.4 100.3 108.3L38.8 5.1zM64 224c0 47.4 18.9 91.2 51.1 126.6L70.7 437.3c-4.9 9.6-4.2 21.1 1.8 30S89.9 480 100.6 480c52.5 0 99.6-19 135.2-41.4c26.2 9.9 54.6 15.4 84.2 15.4c49.6 0 95.8-12.5 134.8-34L98.6 140.4C76.6 165.4 64 193.9 64 224z";
	private const string FA_Language = "M0 128C0 92.7 28.7 64 64 64H576c35.3 0 64 28.7 64 64V384c0 35.3-28.7 64-64 64H64c-35.3 0-64-28.7-64-64V128zM176 160c-13.3 0-24 10.7-24 24v16H88c-13.3 0-24 10.7-24 24s10.7 24 24 24h97.5c-11.5 23.8-28.1 45.4-48.4 63.3c-11.9-11.9-21.5-25.2-28.7-39.6c-5.9-11.9-20.3-16.7-32.2-10.7s-16.7 20.3-10.7 32.2c9.2 18.4 21.2 35.4 35.6 50.4c-11.2 7.8-23.1 14.5-35.6 19.9c-12.2 5.2-17.8 19.3-12.6 31.5s19.3 17.8 31.5 12.6c19.9-8.5 38.6-19.8 55.6-33.5c17 13.7 35.7 25 55.6 33.5c12.2 5.2 26.3-.4 31.5-12.6s-.4-26.3-12.6-31.5c-12.5-5.4-24.4-12.1-35.6-19.9c29.9-26.8 53.6-60.4 68.5-98.6H264c13.3 0 24-10.7 24-24s-10.7-24-24-24H200V184c0-13.3-10.7-24-24-24zm282.4 13.3c-4.1-8.2-12.5-13.3-21.6-13.3s-17.5 5.1-21.6 13.3l-96 192c-5.9 11.9-1.1 26.3 10.7 32.2s26.3 1.1 32.2-10.7L386.8 336h99.9l24.7 49.3c5.9 11.9 20.3 16.7 32.2 10.7s16.7-20.3 10.7-32.2l-96-192zM410.8 288L436.8 236l26 52H410.8z";
	private const string FA_RotateLeft = "M48.5 224H40c-13.3 0-24-10.7-24-24V72c0-9.7 5.8-18.5 14.8-22.2s19.3-1.7 26.2 5.2L98.6 96.6c87.6-86.5 228.7-86.2 315.8 1c87.5 87.5 87.5 229.3 0 316.8s-229.3 87.5-316.8 0c-12.5-12.5-12.5-32.8 0-45.3s32.8-12.5 45.3 0c62.5 62.5 163.8 62.5 226.3 0s62.5-163.8 0-226.3c-62.2-62.2-162.7-62.5-225.3-1L185 183c6.9 6.9 8.9 17.2 5.2 26.2s-12.5 14.8-22.2 14.8H48.5z";
	private const string FA_BuildingColumns = "M243.4 2.6l-224 96c-14 6-21.8 21-18.7 35.8S16.8 160 32 160v8c0 13.3 10.7 24 24 24H456c13.3 0 24-10.7 24-24v-8c15.2 0 28.3-10.7 31.3-25.6s-4.8-29.9-18.7-35.8l-224-96c-8-3.4-17.2-3.4-25.2 0zM128 224H64V420.3c-.6 .3-1.2 .7-1.8 1.1l-48 32c-11.7 7.8-17 22.4-12.9 35.9S17.9 512 32 512H480c14.1 0 26.5-9.2 30.6-22.7s-1.1-28.1-12.9-35.9l-48-32c-.6-.4-1.2-.7-1.8-1.1V224H384V416H344V224H280V416H232V224H168V416H128V224zM256 64a32 32 0 1 1 0 64 32 32 0 1 1 0-64z";
	private const string FA_AlignCenter = "M448 64c0-17.7-14.3-32-32-32H32C14.3 32 0 46.3 0 64S14.3 96 32 96H416c17.7 0 32-14.3 32-32zm0 256c0-17.7-14.3-32-32-32H32c-17.7 0-32 14.3-32 32s14.3 32 32 32H416c17.7 0 32-14.3 32-32zM0 192c0 17.7 14.3 32 32 32H352c17.7 0 32-14.3 32-32s-14.3-32-32-32H32c-17.7 0-32 14.3-32 32zM384 448c0-17.7-14.3-32-32-32H96c-17.7 0-32 14.3-32 32s14.3 32 32 32H352c17.7 0 32-14.3 32-32z";
	private const string FA_AlignJustify = "M448 64c0-17.7-14.3-32-32-32H32C14.3 32 0 46.3 0 64S14.3 96 32 96H416c17.7 0 32-14.3 32-32zm0 128c0-17.7-14.3-32-32-32H32c-17.7 0-32 14.3-32 32s14.3 32 32 32H416c17.7 0 32-14.3 32-32zm0 128c0-17.7-14.3-32-32-32H32c-17.7 0-32 14.3-32 32s14.3 32 32 32H416c17.7 0 32-14.3 32-32zm0 128c0-17.7-14.3-32-32-32H32c-17.7 0-32 14.3-32 32s14.3 32 32 32H416c17.7 0 32-14.3 32-32z";
	private const string FA_PenNib = "M368.4 18.3L312.7 74.1 437.9 199.3l55.7-55.7c21.9-21.9 21.9-57.3 0-79.2L447.6 18.3c-21.9-21.9-57.3-21.9-79.2 0zM288 94.6l-9.2 2.8L134.7 140.6c-19.9 6-35.7 21.2-42.3 41L3.8 445.8c-3.8 11.3-1 23.9 7.3 32.4L164.7 324.7c-3-6.3-4.7-13.3-4.7-20.7c0-26.5 21.5-48 48-48s48 21.5 48 48s-21.5 48-48 48c-7.4 0-14.4-1.7-20.7-4.7L33.7 500.9c8.6 8.3 21.1 11.2 32.4 7.3l264.3-88.6c19.7-6.6 35-22.4 41-42.3l43.2-144.1 2.8-9.2L288 94.6z";
	private const string FA_Expand = "M32 32C14.3 32 0 46.3 0 64v96c0 17.7 14.3 32 32 32s32-14.3 32-32V96h64c17.7 0 32-14.3 32-32s-14.3-32-32-32H32zM64 352c0-17.7-14.3-32-32-32s-32 14.3-32 32v96c0 17.7 14.3 32 32 32h96c17.7 0 32-14.3 32-32s-14.3-32-32-32H64V352zM320 32c-17.7 0-32 14.3-32 32s14.3 32 32 32h64v64c0 17.7 14.3 32 32 32s32-14.3 32-32V64c0-17.7-14.3-32-32-32H320zM448 352c0-17.7-14.3-32-32-32s-32 14.3-32 32v64H320c-17.7 0-32 14.3-32 32s14.3 32 32 32h96c17.7 0 32-14.3 32-32V352z";
	private const string FA_RulerCombined = "M160 288h-56c-13.3 0-24-10.7-24-24s10.7-24 24-24h56V192h-56c-13.3 0-24-10.7-24-24s10.7-24 24-24h56V96h-56c-13.3 0-24-10.7-24-24s10.7-24 24-24h56V32C160 14.3 145.7 0 128 0H32C14.3 0 0 14.3 0 32V480c0 17.7 14.3 32 32 32H480c17.7 0 32-14.3 32-32V384c0-17.7-14.3-32-32-32H464v56c0 13.3-10.7 24-24 24s-24-10.7-24-24V352H368v56c0 13.3-10.7 24-24 24s-24-10.7-24-24V352H272v56c0 13.3-10.7 24-24 24s-24-10.7-24-24V352H160V288z";
	private const string FA_ArrowsLeftRight = "M470.6 278.6c12.5-12.5 12.5-32.8 0-45.3l-104-104c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3L370.7 224 141.3 224l49.4-49.4c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0l-104 104c-12.5 12.5-12.5 32.8 0 45.3l104 104c12.5 12.5 32.8 12.5 45.3 0s12.5-32.8 0-45.3L141.3 288l229.5 0-49.4 49.4c-12.5 12.5-12.5 32.8 0 45.3s32.8 12.5 45.3 0l104-104z";
	private const string FA_ArrowsUpDown = "M278.6 41.4c-12.5-12.5-32.8-12.5-45.3 0l-104 104c-12.5 12.5-12.5 32.8 0 45.3s32.8 12.5 45.3 0L224 141.3V370.7l-49.4-49.4c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3l104 104c12.5 12.5 32.8 12.5 45.3 0l104-104c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0L288 370.7V141.3l49.4 49.4c12.5 12.5 32.8 12.5 45.3 0s12.5-32.8 0-45.3l-104-104z";
	private const string FA_TrashCan = "M135.2 17.7C140.6 6.8 151.7 0 163.8 0H284.2c12.1 0 23.2 6.8 28.6 17.7L320 32h96c17.7 0 32 14.3 32 32s-14.3 32-32 32H32C14.3 96 0 81.7 0 64S14.3 32 32 32h96l7.2-14.3zM32 128H416V448c0 35.3-28.7 64-64 64H96c-35.3 0-64-28.7-64-64V128zm96 64c-8.8 0-16 7.2-16 16V432c0 8.8 7.2 16 16 16s16-7.2 16-16V208c0-8.8-7.2-16-16-16zm96 0c-8.8 0-16 7.2-16 16V432c0 8.8 7.2 16 16 16s16-7.2 16-16V208c0-8.8-7.2-16-16-16zm96 0c-8.8 0-16 7.2-16 16V432c0 8.8 7.2 16 16 16s16-7.2 16-16V208c0-8.8-7.2-16-16-16z";
	private const string FA_Hashtag = "M181.3 32.4c17.4 2.9 29.2 19.4 26.3 36.8L197.8 128h95.1l11.5-69.3c2.9-17.4 19.4-29.2 36.8-26.3s29.2 19.4 26.3 36.8L357.8 128H416c17.7 0 32 14.3 32 32s-14.3 32-32 32H347.1L325.8 320H384c17.7 0 32 14.3 32 32s-14.3 32-32 32H315.1l-11.5 69.3c-2.9 17.4-19.4 29.2-36.8 26.3s-29.2-19.4-26.3-36.8l9.8-58.7H155.1l-11.5 69.3c-2.9 17.4-19.4 29.2-36.8 26.3s-29.2-19.4-26.3-36.8L90.2 384H32c-17.7 0-32-14.3-32-32s14.3-32 32-32h68.9l21.3-128H64c-17.7 0-32-14.3-32-32s14.3-32 32-32h68.9l11.5-69.3c2.9-17.4 19.4-29.2 36.8-26.3zM187.1 192L165.8 320h95.1l21.3-128H187.1z";
	private const string FA_LayerGroup = "M264.5 5.2c14.9-6.9 32.1-6.9 47 0l218.6 101c8.5 3.9 13.9 12.4 13.9 21.8s-5.4 17.9-13.9 21.8l-218.6 101c-14.9 6.9-32.1 6.9-47 0L45.9 149.8C37.4 145.8 32 137.3 32 128s5.4-17.9 13.9-21.8L264.5 5.2zM476.9 209.6l53.2 24.6c8.5 3.9 13.9 12.4 13.9 21.8s-5.4 17.9-13.9 21.8l-218.6 101c-14.9 6.9-32.1 6.9-47 0L45.9 277.8C37.4 273.8 32 265.3 32 256s5.4-17.9 13.9-21.8l53.2-24.6 152 70.2c23.4 10.8 50.4 10.8 73.8 0l152-70.2zm-152 198.2l152-70.2 53.2 24.6c8.5 3.9 13.9 12.4 13.9 21.8s-5.4 17.9-13.9 21.8l-218.6 101c-14.9 6.9-32.1 6.9-47 0L45.9 405.8C37.4 401.8 32 393.3 32 384s5.4-17.9 13.9-21.8l53.2-24.6 152 70.2c23.4 10.8 50.4 10.8 73.8 0z";
	private const string FA_Bolt = "M349.4 44.6c5.9-13.7 1.5-29.7-10.6-38.5s-28.6-8-39.9 1.8l-256 224c-10 8.8-13.6 22.9-8.9 35.3S50.7 288 64 288H175.5L98.6 467.4c-5.9 13.7-1.5 29.7 10.6 38.5s28.6 8 39.9-1.8l256-224c10-8.8 13.6-22.9 8.9-35.3s-16.6-20.7-30-20.7H272.5L349.4 44.6z";
	private const string FA_Star = "M316.9 18C311.6 7 300.4 0 288.1 0s-23.4 7-28.8 18L195 150.3 51.4 171.5c-12 1.8-22 10.2-25.7 21.7s-.7 24.2 7.9 32.7L137.8 329 113.2 474.7c-2 12 3 24.2 12.9 31.3s23 8 33.8 2.3l128.3-68.5 128.3 68.5c10.8 5.7 23.9 4.9 33.8-2.3s14.9-19.3 12.9-31.3L438.5 329 542.7 225.9c8.6-8.5 11.7-21.2 7.9-32.7s-13.7-19.9-25.7-21.7L381.2 150.3 316.9 18z";
	private const string FA_Flag = "M64 32C64 14.3 49.7 0 32 0S0 14.3 0 32V64 368 480c0 17.7 14.3 32 32 32s32-14.3 32-32V352l64.3-16.1c41.1-10.3 84.6-5.5 122.5 13.4c44.2 22.1 95.5 24.8 141.7 7.4l34.7-13c12.5-4.7 20.8-16.6 20.8-30V66.1c0-23-24.2-38-44.8-27.7l-9.6 4.8c-46.3 23.2-100.8 23.2-147.1 0c-35.1-17.6-75.4-22-113.5-12.5L64 48V32z";
	private const string FA_ShieldHalved = "M256 0c4.6 0 9.2 1 13.4 2.9L457.7 82.8c22 9.3 38.4 31 38.3 57.2c-.5 99.2-41.3 280.7-213.6 363.2c-16.7 8-36.1 8-52.8 0C57.3 420.7 16.5 239.2 16 140c-.1-26.2 16.3-47.9 38.3-57.2L242.7 2.9C246.8 1 251.4 0 256 0zm0 66.8V444.8C394 378 431.1 230.1 432 141.4L256 66.8l0 0z";
	private const string FA_Heading = "M32 64C14.3 64 0 78.3 0 96s14.3 32 32 32H64V224 384H32c-17.7 0-32 14.3-32 32s14.3 32 32 32H160c17.7 0 32-14.3 32-32s-14.3-32-32-32H128V288H320v96H288c-17.7 0-32 14.3-32 32s14.3 32 32 32H416c17.7 0 32-14.3 32-32s-14.3-32-32-32H384V224 128h32c17.7 0 32-14.3 32-32s-14.3-32-32-32H288c-17.7 0-32 14.3-32 32s14.3 32 32 32h32v96H128V128h32c17.7 0 32-14.3 32-32s-14.3-32-32-32H32z";
	private const string FA_PaperPlane = "M498.1 5.6c10.1 7 15.4 19.1 13.5 31.2l-64 416c-1.5 9.7-7.4 18.2-16 23s-18.9 5.4-28 1.6L284 427.7l-68.5 74.1c-8.9 9.7-22.9 12.9-35.2 8.1S160 493.2 160 480V396.4c0-4 1.5-7.8 4.2-10.7L331.8 202.8c5.8-6.3 5.6-16-.4-22s-15.7-6.4-22-.7L106 360.8 17.7 316.6C7.1 311.3 .3 300.7 0 288.9s5.9-22.8 16.1-28.7l448-256c10.7-6.1 23.9-5.5 34 1.4z";
	private const string FA_Lock = "M144 144v48H304V144c0-44.2-35.8-80-80-80s-80 35.8-80 80zM80 192V144C80 64.5 144.5 0 224 0s144 64.5 144 144v48h16c35.3 0 64 28.7 64 64V448c0 35.3-28.7 64-64 64H64c-35.3 0-64-28.7-64-64V256c0-35.3 28.7-64 64-64H80z";
	private const string FA_ScrewdriverWrench = "M78.6 5C69.1-2.4 55.6-1.5 47 7L7 47c-8.5 8.5-9.4 22-2.1 31.6l80 104c4.5 5.9 11.6 9.4 19 9.4h54.1l109 109c-14.7 29-10 65.4 14.3 89.6l112 112c12.5 12.5 32.8 12.5 45.3 0l64-64c12.5-12.5 12.5-32.8 0-45.3l-112-112c-24.2-24.2-60.6-29-89.6-14.3l-109-109V104c0-7.5-3.5-14.5-9.4-19L78.6 5zM19.9 396.1C7.2 408.8 0 426.1 0 444.1C0 481.6 30.4 512 67.9 512c18 0 35.3-7.2 48-19.9L233.7 374.3c-7.8-20.9-9-43.6-3.6-65.1l-61.7-61.7L19.9 396.1zM512 144c0-10.5-1.1-20.7-3.2-30.5c-2.4-11.2-16.1-14.1-24.2-6l-63.9 63.9c-3 3-7.1 4.7-11.3 4.7H352c-8.8 0-16-7.2-16-16V102.6c0-4.2 1.7-8.3 4.7-11.3l63.9-63.9c8.1-8.1 5.2-21.8-6-24.2C388.7 1.1 378.5 0 368 0C288.5 0 224 64.5 224 144l0 .8 85.3 85.3c36-9.1 75.8 .5 104 28.7L429 274.5c49-23 83-72.8 83-130.5zM56 432a24 24 0 1 1 48 0 24 24 0 1 1 -48 0z";
	private const string FA_TableCells = "M64 32C28.7 32 0 60.7 0 96V416c0 35.3 28.7 64 64 64H448c35.3 0 64-28.7 64-64V96c0-35.3-28.7-64-64-64H64zm88 64v64H64V96h88zm56 0h88v64H208V96zm240 0v64H360V96h88zM64 224h88v64H64V224zm232 0v64H208V224h88zm64 0h88v64H360V224zM152 352v64H64V352h88zm56 0h88v64H208V352zm240 0v64H360V352h88z";
	private const string FA_Broom = "M568.4 39.2c10.2 10.2 10.2 26.8 0 37.1L367.5 277.2l-32.9-32.9L535.5 43.4c10.2-10.2 26.8-10.2 37.1 0zM301.7 211.4l-24.8-24.8c-15.6-15.6-40.9-15.6-56.6 0L192 214.9 361.1 384l28.3-28.3c15.6-15.6 15.6-40.9 0-56.6l-24.8-24.8-62.9-62.9zM169.4 237.5L13.1 393.8C4.7 402.2 0 413.6 0 425.5C0 450.1 19.9 470 44.5 470c9.6 0 18.9-3.1 26.6-8.8l46.9-35.2c6.6-5 15.9-4.3 21.7 1.5s6.5 15.1 1.5 21.7l-35.2 46.9c-5.7 7.7-8.8 17-8.8 26.6C97.2 547.3 117.1 567.2 141.7 567.2c11.9 0 23.3-4.7 31.7-13.1L329.7 397.8 169.4 237.5z";
	private const string FA_Eraser = "M290.7 57.4L57.4 290.7c-25 25-25 65.5 0 90.5l80 80c12 12 28.3 18.7 45.3 18.7H288h9.4H512c17.7 0 32-14.3 32-32s-14.3-32-32-32H387.9L518.6 285.3c25-25 25-65.5 0-90.5L381.3 57.4c-25-25-65.5-25-90.5 0zM297.4 416H288l-105.4 0-80-80L227.3 211.3 364.7 348.7 297.4 416z";
	private const string FA_Font = "M254 52.8C249.3 40.3 237.3 32 224 32s-25.3 8.3-30 20.8L57.8 416H32c-17.7 0-32 14.3-32 32s14.3 32 32 32h96c17.7 0 32-14.3 32-32s-14.3-32-32-32h-1.8l18-48H303.8l18 48H320c-17.7 0-32 14.3-32 32s14.3 32 32 32h96c17.7 0 32-14.3 32-32s-14.3-32-32-32H390.2L254 52.8zM279.8 304H168.2L224 155.1 279.8 304z";
	private const string FA_Coins = "M512 80c0 18-14.3 34.6-38.4 48c-29.1 16.1-72.5 27.5-122.3 30.9c-3.7-1.8-7.4-3.5-11.3-5C300.6 137.4 248.2 128 192 128c-8.3 0-16.4 .2-24.5 .6l-1.1-.6C142.3 114.6 128 98 128 80c0-44.2 86-80 192-80S512 35.8 512 80zM160.7 161.1c10.2-.7 20.7-1.1 31.3-1.1c62.2 0 117.4 12.3 152.5 31.4C369.3 204.9 384 221.7 384 240c0 4-.7 7.9-2.1 11.7c-4.6 13.2-17 25.3-35 35.5c0 0 0 0 0 0c-.1 .1-.3 .1-.4 .2l0 0 0 0c-.3 .2-.6 .3-.9 .5c-35 19.4-90.8 32-153.6 32c-59.6 0-112.9-11.3-148.2-29.1c-1.9-.9-3.7-1.9-5.5-2.9C14.3 274.6 0 258 0 240c0-34.8 53.4-64.5 128-75.4c10.5-1.5 21.4-2.7 32.7-3.5zM384 336c0 18-14.3 34.6-38.4 48c-35.3 19.6-91.4 32-153.6 32s-118.3-12.4-153.6-32C14.3 370.6 0 354 0 336V296.6c14.5 10.5 34 19.7 57.6 27c40.3 12.6 88 19.4 134.4 19.4s94.1-6.8 134.4-19.4c23.6-7.3 43.1-16.4 57.6-27V336zm0 96c0 18-14.3 34.6-38.4 48c-35.3 19.6-91.4 32-153.6 32s-118.3-12.4-153.6-32C14.3 466.6 0 450 0 432V392.6c14.5 10.5 34 19.7 57.6 27c40.3 12.6 88 19.4 134.4 19.4s94.1-6.8 134.4-19.4c23.6-7.3 43.1-16.4 57.6-27V432z";
	private const string FA_TableColumns = "M0 96C0 60.7 28.7 32 64 32H448c35.3 0 64 28.7 64 64V416c0 35.3-28.7 64-64 64H64c-35.3 0-64-28.7-64-64V96zm64 64V416H224V160H64zm384 0H288V416H448V160z";
	private const string FA_BookOpen = "M249.6 471.5c10.8 3.8 22.4-4.1 22.4-15.5V78.6c0-4.2-1.6-8.4-5-11C247.4 52 202.4 32 144 32C93.5 32 46.3 45.3 18.1 56.1C6.8 60.5 0 71.7 0 83.8V454.1c0 11.9 12.8 20.2 24.1 16.5C55.6 460.1 105.5 448 144 448c33.9 0 79 14 105.6 23.5zm76.8 0C353 462 398.1 448 432 448c38.5 0 88.4 12.1 119.9 22.6c11.3 3.8 24.1-4.6 24.1-16.5V83.8c0-12.1-6.8-23.3-18.1-27.6C529.7 45.3 482.5 32 432 32c-58.4 0-103.4 20-123 35.6c-3.3 2.6-5 6.8-5 11V456c0 11.4 11.7 19.3 22.4 15.5z";
	private const string FA_FolderOpen = "M64 480H448c35.3 0 64-28.7 64-64V160c0-35.3-28.7-64-64-64H288c-10.1 0-19.6-4.7-25.6-12.8L243.2 57.6C231.1 41.5 212.1 32 192 32H64C28.7 32 0 60.7 0 96V416c0 35.3 28.7 64 64 64z";
	private const string FA_FileLines = "M64 464c-8.8 0-16-7.2-16-16V64c0-8.8 7.2-16 16-16H224v80c0 17.7 14.3 32 32 32h80V448c0 8.8-7.2 16-16 16H64zM64 0C28.7 0 0 28.7 0 64V448c0 35.3 28.7 64 64 64H320c35.3 0 64-28.7 64-64V154.5c0-17-6.7-33.3-18.7-45.3L274.7 18.7C262.7 6.7 246.5 0 229.5 0H64zm56 256c-13.3 0-24 10.7-24 24s10.7 24 24 24H264c13.3 0 24-10.7 24-24s-10.7-24-24-24H120zm0 96c-13.3 0-24 10.7-24 24s10.7 24 24 24H264c13.3 0 24-10.7 24-24s-10.7-24-24-24H120z";
	private const string FA_FileCirclePlus = "M0 64C0 28.7 28.7 0 64 0H224V128c0 17.7 14.3 32 32 32H384v38.6C310.1 219.5 256 287.4 256 368c0 59.1 29.1 111.3 73.7 143.3c-3.2 .5-6.4 .7-9.7 .7H64c-35.3 0-64-28.7-64-64V64zm384 64H256V0L384 128zm48 96a144 144 0 1 1 0 288 144 144 0 1 1 0-288zm16 80c0-8.8-7.2-16-16-16s-16 7.2-16 16v48H368c-8.8 0-16 7.2-16 16s7.2 16 16 16h48v48c0 8.8 7.2 16 16 16s16-7.2 16-16V384h48c8.8 0 16-7.2 16-16s-7.2-16-16-16H448V304z";
	private const string FA_Award = "M173.8 5.5c11-7.3 25.4-7.3 36.4 0L228 17.2c6 3.9 13 5.8 20.1 5.4l21.3-1.3c13.2-.8 25.6 6.4 31.5 18.2l9.6 19.1c3.2 6.4 8.4 11.5 14.7 14.7L344.5 83c11.8 5.9 19 18.3 18.2 31.5l-1.3 21.3c-.4 7.1 1.5 14.2 5.4 20.1l11.8 17.8c7.3 11 7.3 25.4 0 36.4L366.8 228c-3.9 6-5.8 13-5.4 20.1l1.3 21.3c.8 13.2-6.4 25.6-18.2 31.5l-19.1 9.6c-6.4 3.2-11.5 8.4-14.7 14.7L301 344.5c-5.9 11.8-18.3 19-31.5 18.2l-21.3-1.3c-7.1-.4-14.2 1.5-20.1 5.4l-17.8 11.8c-11 7.3-25.4 7.3-36.4 0L156 366.8c-6-3.9-13-5.8-20.1-5.4l-21.3 1.3c-13.2 .8-25.6-6.4-31.5-18.2l-9.6-19.1c-3.2-6.4-8.4-11.5-14.7-14.7L39.5 301c-11.8-5.9-19-18.3-18.2-31.5l1.3-21.3c.4-7.1-1.5-14.2-5.4-20.1L5.5 210.2c-7.3-11-7.3-25.4 0-36.4L17.2 156c3.9-6 5.8-13 5.4-20.1l-1.3-21.3c-.8-13.2 6.4-25.6 18.2-31.5l19.1-9.6C65 70.2 70.2 65 73.4 58.6L83 39.5c5.9-11.8 18.3-19 31.5-18.2l21.3 1.3c7.1 .4 14.2-1.5 20.1-5.4L173.8 5.5zM272 192a80 80 0 1 0 -160 0 80 80 0 1 0 160 0zM1.3 441.8L44.4 339.3c.2 .1 .3 .2 .4 .4l9.6 19.1c11.7 23.2 36 37.3 62 35.8l21.3-1.3c.2 0 .5 0 .7 .2l17.8 11.8c5.1 3.3 10.5 5.9 16.1 7.7l-37.6 89.3c-2.3 5.5-7.4 9.2-13.3 9.7s-11.6-2.2-14.8-7.2L74.4 455.5l-56.1 8.3c-5.7 .8-11.4-1.5-15-6s-4.3-10.7-2-16zm248 60.4L211.7 413c5.6-1.8 11-4.3 16.1-7.7l17.8-11.8c.2-.1 .4-.2 .7-.2l21.3 1.3c26 1.5 50.3-12.6 62-35.8l9.6-19.1c.1-.2 .2-.3 .4-.4l43.2 102.5c2.3 5.3 1.5 11.4-2 16s-9.3 6.9-15 6l-56.1-8.3-32.2 49.2c-3.2 5-8.9 7.7-14.8 7.2s-11-4.3-13.3-9.7z";
	private const string FA_GraduationCap = "M320 32c-8.1 0-16.1 1.4-23.7 4.1L15.8 137.4C6.3 140.9 0 149.9 0 160s6.3 19.1 15.8 22.6l57.9 20.9C57.3 229.3 48 259.8 48 291.9v28.1c0 28.4-10.8 57.7-22.3 80.8c-6.5 13-13.9 25.8-22.5 37.6C0 442.7-.9 448.3 .9 453.4s6 8.9 11.2 10.2l64 16c4.2 1.1 8.7 .3 12.4-2s6.3-6.1 7.1-10.4c8.6-42.8 4.3-81.2-2.1-108.7C90.3 344.3 86 329.8 80 316.5V291.9c0-30.2 10.2-58.7 27.9-81.5c12.9-15.5 29.6-28 49.2-35.7l157-61.7c8.2-3.2 17.5 .8 20.7 9s-.8 17.5-9 20.7l-157 61.7c-12.4 4.9-23.3 12.4-32.2 21.6l159.6 57.6c7.6 2.7 15.6 4.1 23.7 4.1s16.1-1.4 23.7-4.1L624.2 182.6c9.5-3.4 15.8-12.5 15.8-22.6s-6.3-19.1-15.8-22.6L343.7 36.1C336.1 33.4 328.1 32 320 32zM128 408c0 35.3 86 72 192 72s192-36.7 192-72L496.7 262.6 354.5 314c-11.1 4-22.8 6-34.5 6s-23.5-2-34.5-6L143.3 262.6 128 408z";
	private const string FA_Sliders = "M0 416c0 17.7 14.3 32 32 32l54.7 0c12.3 28.3 40.5 48 73.3 48s61-19.7 73.3-48L480 448c17.7 0 32-14.3 32-32s-14.3-32-32-32l-246.7 0c-12.3-28.3-40.5-48-73.3-48s-61 19.7-73.3 48L32 384c-17.7 0-32 14.3-32 32zm128 0a32 32 0 1 1 64 0 32 32 0 1 1 -64 0zM320 256a32 32 0 1 1 64 0 32 32 0 1 1 -64 0zm32-80c-32.8 0-61 19.7-73.3 48L32 224c-17.7 0-32 14.3-32 32s14.3 32 32 32l246.7 0c12.3 28.3 40.5 48 73.3 48s61-19.7 73.3-48l54.7 0c17.7 0 32-14.3 32-32s-14.3-32-32-32l-54.7 0c-12.3-28.3-40.5-48-73.3-48zM192 128a32 32 0 1 1 0-64 32 32 0 1 1 0 64zm73.3-64C253 35.7 224.8 16 192 16s-61 19.7-73.3 48L32 64C14.3 64 0 78.3 0 96s14.3 32 32 32l86.7 0c12.3 28.3 40.5 48 73.3 48s61-19.7 73.3-48L480 128c17.7 0 32-14.3 32-32s-14.3-32-32-32L265.3 64z";
	private const string FA_CircleInfo = "M256 512A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM216 336h24V272H216c-13.3 0-24-10.7-24-24s10.7-24 24-24h48c13.3 0 24 10.7 24 24v88h8c13.3 0 24 10.7 24 24s-10.7 24-24 24H216c-13.3 0-24-10.7-24-24s10.7-24 24-24zm40-208a32 32 0 1 1 0 64 32 32 0 1 1 0-64z";
	private const string FA_CloudArrowUp = "M0 336c0 79.5 64.5 144 144 144H512c70.7 0 128-57.3 128-128c0-61.9-44-113.6-102.4-125.4c4.1-10.7 6.4-22.4 6.4-34.6c0-53-43-96-96-96c-19.7 0-38.1 6-53.3 16.2C367 64.2 315.3 32 256 32C167.6 32 96 103.6 96 192c0 2.7 .1 5.4 .2 8.1C40.2 219.8 0 273.2 0 336zm393-89c9.4 9.4 9.4 24.6 0 33.9s-24.6 9.4-33.9 0l-39-39V392c0 13.3-10.7 24-24 24s-24-10.7-24-24V241.9l-39 39c-9.4 9.4-24.6 9.4-33.9 0s-9.4-24.6 0-33.9l80-80c9.4-9.4 24.6-9.4 33.9 0l80 80z";

	private const string FA_Print = "M128 0C92.7 0 64 28.7 64 64v96h64V64H384v96h64V64c0-35.3-28.7-64-64-64H128zM384 448H128V320H384v128zm64-256H64c-35.3 0-64 28.7-64 64V416c0 35.3 28.7 64 64 64h32v-32H64c-17.7 0-32-14.3-32-32V256c0-17.7 14.3-32 32-32H448c17.7 0 32 14.3 32 32v160c0 17.7-14.3 32-32 32H416v32h32c35.3 0 64-28.7 64-64V256c0-35.3-28.7-64-64-64zM432 296a24 24 0 1 0 0-48 24 24 0 1 0 0 48z";
	private const string FA_FilePdf = "M0 64C0 28.7 28.7 0 64 0H224V128c0 17.7 14.3 32 32 32H384V448c0 35.3-28.7 64-64 64H64c-35.3 0-64-28.7-64-64V64zm384 64H256V0L384 128zM176 352h32c30.9 0 56-25.1 56-56s-25.1-56-56-56H152c-13.3 0-24 10.7-24 24v128c0 13.3 10.7 24 24 24s24-10.7 24-24V352zm24-64c4.4 0 8 3.6 8 8s-3.6 8-8 8h-24V288h24z";
	private const string FA_Qrcode = "M0 96C0 43 43 0 96 0h96c53 0 96 43 96 96v96c0 53-43 96-96 96H96c-53 0-96-43-96-96V96zm64 0v96c0 17.7 14.3 32 32 32h96c17.7 0 32-14.3 32-32V96c0-17.7-14.3-32-32-32H96c-17.7 0-32 14.3-32 32zM288 96c0-53 43-96 96-96h96c53 0 96 43 96 96v96c0 53-43 96-96 96H384c-53 0-96-43-96-96V96zm64 0v96c0 17.7 14.3 32 32 32h96c17.7 0 32-14.3 32-32V96c0-17.7-14.3-32-32-32H384c-17.7 0-32 14.3-32 32zM0 384c0-53 43-96 96-96h96c53 0 96 43 96 96v96c0 53-43 96-96 96H96c-53 0-96-43-96-96V384zm64 0v96c0 17.7 14.3 32 32 32h96c17.7 0 32-14.3 32-32V384c0-17.7-14.3-32-32-32H96c-17.7 0-32 14.3-32 32zm352-64c0-17.7 14.3-32 32-32h64c17.7 0 32 14.3 32 32v64c0 17.7-14.3 32-32 32H448c-17.7 0-32-14.3-32-32V320zm-96 96c0-17.7 14.3-32 32-32h64c17.7 0 32 14.3 32 32v64c0 17.7-14.3 32-32 32H352c-17.7 0-32-14.3-32-32V416z";

	// Biểu tượng thương hiệu nhận diện tác giả Hà Xuân Vũ (H - V Monogram cách tân)
	private const string FA_HV_Brand = "M96 64C87 64 80 71 80 80L80 275C80 285 85 294 93 300L244 457C251 464 261 464 268 457L419 300C427 294 432 285 432 275L432 80C432 71 425 64 416 64L368 64C359 64 352 71 352 80L352 192L160 192L160 80C160 71 153 64 144 64ZM160 256L352 256L256 376Z";
	// Biểu tượng liên hệ số điện thoại
	private const string FA_Phone = "M164.9 24.6c-7.7-18.6-28-28.5-47.4-23.2l-88 24C12.1 30.2 0 46 0 64C0 311.4 200.6 512 448 512c18 0 33.8-12.1 38.6-29.5l24-88c5.3-19.4-4.6-39.7-23.2-47.4l-96-40c-16.3-6.8-35.2-2.1-46.3 11.6L304.7 368C234.3 334.7 177.3 277.7 144 207.3L193.3 167c13.7-11.2 18.4-30 11.6-46.3l-40-96z";

	private static readonly Dictionary<string, Bitmap> _cache = new Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);
	private static readonly object _lock = new object();

	public static Bitmap GetIcon(string controlId)
	{
		IconSpec spec = ResolveSpec(controlId ?? "");
		string cacheKey = ((int)spec.Theme) + "|" + spec.PathData.GetHashCode();
		lock (_lock)
		{
			if (_cache.TryGetValue(cacheKey, out var cached) && cached != null)
			{
				return cached;
			}
			Bitmap rendered = RenderSpec(spec);
			if (rendered != null)
			{
				_cache[cacheKey] = rendered;
			}
			return rendered;
		}
	}

	private static IconSpec ResolveSpec(string id)
	{
		switch (id)
		{
		// Large Ribbon Buttons (3D Glossy Coral / Amber / Terracotta Badges)
		case "mnuQuickInsert":
			return new IconSpec(FA_LayerGroup, IconTheme.HeroCoral);
		case "mnuTools":
			return new IconSpec(FA_ScrewdriverWrench, IconTheme.HeroAmber);
		case "btnToggleTaskPane":
			return new IconSpec(FA_TableColumns, IconTheme.HeroTerracotta);
		case "btnAutoStandardize":
			return new IconSpec(FA_WandMagicSparkles, IconTheme.HeroCoral);
		case "btnSpellCheck":
			return new IconSpec(FA_SpellCheck, IconTheme.HeroAmber);
		case "btnPageSetupA4":
			return new IconSpec(FA_Expand, IconTheme.HeroCoral);
		case "btnPageNum":
			return new IconSpec(FA_Hashtag, IconTheme.HeroAmber);
		case "mnuDocTemplates":
			return new IconSpec(FA_BookOpen, IconTheme.HeroCoral);
		case "btnPersonalize":
			return new IconSpec(FA_HV_Brand, IconTheme.HeroCoral);
		case "btnContactAuthor":
			return new IconSpec(FA_Phone, IconTheme.HeroAmber);
		case "btnLatPdf":
			return new IconSpec(FA_FilePdf, IconTheme.HeroCoral);

		// Group: Chuẩn hóa & Quy chuẩn
		case "btnDetectDoc":
			return new IconSpec(FA_MagnifyingGlass, IconTheme.CoralDuotone);
		case "btnAuditDoc":
			return new IconSpec(FA_ClipboardCheck, IconTheme.EmeraldDuotone);
		case "mnuDeleteComments":
		case "btnDeleteAuditComments":
		case "btnDeleteAllComments":
			return new IconSpec(FA_CommentSlash, IconTheme.CrimsonDuotone);
		case "btnUnicodeConvert":
			return new IconSpec(FA_Language, IconTheme.AmberDuotone);
		case "btnUndoStandardize":
			return new IconSpec(FA_RotateLeft, IconTheme.CoralDuotone);

		// Group: Chuẩn hóa Thủ công
		case "mnuManualStandardize":
			return new IconSpec(FA_AlignJustify, IconTheme.HeroCoral);
		case "btnManualHeader":
			return new IconSpec(FA_BuildingColumns, IconTheme.CoralDuotone);
		case "btnManualTitle":
			return new IconSpec(FA_AlignCenter, IconTheme.AmberDuotone);
		case "btnManualBody":
			return new IconSpec(FA_AlignJustify, IconTheme.CoralDuotone);
		case "btnManualHeadings":
			return new IconSpec(FA_Heading, IconTheme.EmeraldDuotone);
		case "btnManualSignature":
			return new IconSpec(FA_PenNib, IconTheme.AmberDuotone);
		case "btnManualSpacingCheck":
			return new IconSpec(FA_ArrowsLeftRight, IconTheme.CoralDuotone);

		// Group: Bố cục Trang
		case "btnOpenPageMarginDlg":
			return new IconSpec(FA_RulerCombined, IconTheme.CoralDuotone);
		case "btnLandscapePage":
			return new IconSpec(FA_ArrowsLeftRight, IconTheme.AmberDuotone);
		case "btnPortraitPage":
			return new IconSpec(FA_ArrowsUpDown, IconTheme.CoralDuotone);
		case "btnDeleteBlankPages":
			return new IconSpec(FA_TrashCan, IconTheme.CrimsonDuotone);

		// In an
		case "mnuPrintOptions":
		case "btnPrintDuplexDialog":
			return new IconSpec(FA_Print, IconTheme.HeroCoral);
		case "btnPrintOddPages":
			return new IconSpec(FA_Print, IconTheme.AmberDuotone);
		case "btnPrintEvenPages":
			return new IconSpec(FA_Print, IconTheme.EmeraldDuotone);

		// PDF
		case "mnuPdfTools":
			return new IconSpec(FA_FilePdf, IconTheme.HeroAmber);
		case "btnConvertPdfToWord":
		case "btnBatchPdfToWord":
			return new IconSpec(FA_FilePdf, IconTheme.CoralDuotone);
		case "btnExportCurrentToPdf":
			return new IconSpec(FA_FilePdf, IconTheme.CrimsonDuotone);

		// QR Code
		case "mnuQrTools":
		case "btnOpenQrManager":
			return new IconSpec(FA_Qrcode, IconTheme.HeroTerracotta);
		case "btnCreateDocQr":
			return new IconSpec(FA_BuildingColumns, IconTheme.CoralDuotone);
		case "btnCreateSelectionQr":
			return new IconSpec(FA_Qrcode, IconTheme.AmberDuotone);
		case "btnCreateVietQr":
			return new IconSpec(FA_Coins, IconTheme.EmeraldDuotone);

		// Group: Công cụ
		case "btnFormatTables":
			return new IconSpec(FA_TableCells, IconTheme.CoralDuotone);
		case "btnCleanSpaces":
			return new IconSpec(FA_Broom, IconTheme.AmberDuotone);
		case "btnRemoveToneSelection":
			return new IconSpec(FA_Eraser, IconTheme.CrimsonDuotone);
		case "btnFixToneMarks":
			return new IconSpec(FA_SpellCheck, IconTheme.EmeraldDuotone);
		case "btnFixToneMarksOld":
			return new IconSpec(FA_Font, IconTheme.CoralDuotone);
		case "btnNumberToVndWords":
			return new IconSpec(FA_Coins, IconTheme.AmberDuotone);

		// Group: Biểu mẫu & Cài đặt
		case "btnOpenUnifiedTemplateLibrary":
			return new IconSpec(FA_BookOpen, IconTheme.CoralDuotone);
		case "mnuMyTemplates":
		case "dynMyTemplates":
		case "btnEmptyUserTpl":
			return new IconSpec(FA_FolderOpen, IconTheme.AmberDuotone);
		case "btnOpenMyTemplatesMgr":
		case "btnDynOpenMgr":
		case "btnOpenPersonalizeInMenu":
			return new IconSpec(FA_Sliders, IconTheme.CoralDuotone);
		case "btnAddUserTpl":
		case "btnDynAddUserTpl":
			return new IconSpec(FA_FileCirclePlus, IconTheme.EmeraldDuotone);
		case "mnuSubTDKT":
			return new IconSpec(FA_Award, IconTheme.AmberDuotone);
		case "mnuSubHanhChinh":
		case "mnuSubTinh":
		case "mnuSubXa":
			return new IconSpec(FA_BuildingColumns, IconTheme.CoralDuotone);
		case "mnuSubBQP":
			return new IconSpec(FA_ShieldHalved, IconTheme.EmeraldDuotone);
		case "mnuSubDang":
		case "mnuSubDangTW":
		case "mnuSubDangChiBo":
			return new IconSpec(FA_Flag, IconTheme.CrimsonDuotone);
		case "mnuSubDoan":
		case "mnuSubDoanChung":
		case "mnuSubDoanDaiHoi":
			return new IconSpec(FA_Star, IconTheme.AmberDuotone);
		case "mnuSubPersonal":
			return new IconSpec(FA_PenNib, IconTheme.CoralDuotone);
		case "mnuSubOther":
			return new IconSpec(FA_GraduationCap, IconTheme.AmberDuotone);
		case "mnuHeaderDateLayout":
			return new IconSpec(FA_TableColumns, IconTheme.CoralDuotone);
		case "btnHeaderDateSeparate":
			return new IconSpec(FA_TableCells, IconTheme.EmeraldDuotone);
		case "btnHeaderDateTopCell":
			return new IconSpec(FA_FileLines, IconTheme.AmberDuotone);
		case "btnAboutApp":
			return new IconSpec(FA_CircleInfo, IconTheme.CoralDuotone);
		case "btnCheckUpdate":
			return new IconSpec(FA_CloudArrowUp, IconTheme.AmberDuotone);
		}

		// Prefix matches for Quick Insert & Templates
		if (id.StartsWith("btnInsQuocHieu", StringComparison.OrdinalIgnoreCase))
		{
			if (id.IndexOf("Auto", StringComparison.OrdinalIgnoreCase) >= 0)
				return new IconSpec(FA_Bolt, IconTheme.AmberDuotone);
			if (id.IndexOf("Dang", StringComparison.OrdinalIgnoreCase) >= 0)
				return new IconSpec(FA_Flag, IconTheme.CrimsonDuotone);
			if (id.IndexOf("Doan", StringComparison.OrdinalIgnoreCase) >= 0)
				return new IconSpec(FA_Star, IconTheme.AmberDuotone);
			if (id.IndexOf("BQP", StringComparison.OrdinalIgnoreCase) >= 0)
				return new IconSpec(FA_ShieldHalved, IconTheme.EmeraldDuotone);
			return new IconSpec(FA_BuildingColumns, IconTheme.CoralDuotone);
		}
		if (id == "mnuSubHeaders" || id == "mnuSubAdminAgencies" || id.StartsWith("btnInsAdmin", StringComparison.OrdinalIgnoreCase) || id == "btnInsFooterAgencyInfo")
		{
			return new IconSpec(FA_BuildingColumns, IconTheme.CoralDuotone);
		}
		if (id == "mnuSubYouthAgencies" || id.StartsWith("btnInsYouth", StringComparison.OrdinalIgnoreCase))
		{
			return new IconSpec(FA_Star, IconTheme.AmberDuotone);
		}
		if (id == "mnuSubDocTitle" || id.StartsWith("btnInsTitle", StringComparison.OrdinalIgnoreCase))
		{
			return new IconSpec(FA_Heading, IconTheme.CoralDuotone);
		}
		if (id == "mnuSubSavedRecipients" || id.StartsWith("btnInsRecSaved", StringComparison.OrdinalIgnoreCase))
		{
			return new IconSpec(FA_PaperPlane, IconTheme.AmberDuotone);
		}
		if (id == "mnuSubSigners" || id.StartsWith("mnuSub", StringComparison.OrdinalIgnoreCase) && id.IndexOf("Signer", StringComparison.OrdinalIgnoreCase) >= 0 || id.StartsWith("btnInsSig", StringComparison.OrdinalIgnoreCase) || id.StartsWith("btnInsPartySig", StringComparison.OrdinalIgnoreCase) || id.StartsWith("btnInsMilitarySig", StringComparison.OrdinalIgnoreCase))
		{
			return new IconSpec(FA_PenNib, IconTheme.CoralDuotone);
		}
		if (id == "mnuSubSpecialMarks" || id.StartsWith("btnInsDoKhan", StringComparison.OrdinalIgnoreCase))
		{
			return new IconSpec(FA_Bolt, IconTheme.CrimsonDuotone);
		}
		if (id.StartsWith("btnInsDoMat", StringComparison.OrdinalIgnoreCase) || id.StartsWith("btnInsPhamVi", StringComparison.OrdinalIgnoreCase))
		{
			return new IconSpec(FA_Lock, IconTheme.AmberDuotone);
		}
		if (id.StartsWith("tpl", StringComparison.OrdinalIgnoreCase) || id.StartsWith("btnUserTpl", StringComparison.OrdinalIgnoreCase))
		{
			return new IconSpec(FA_FileLines, IconTheme.CoralDuotone);
		}

		return new IconSpec(FA_WandMagicSparkles, IconTheme.CoralDuotone);
	}

	private static Bitmap RenderSpec(IconSpec spec)
	{
		try
		{
			const int size = 64;
			DrawingVisual visual = new DrawingVisual();
			using (DrawingContext dc = visual.RenderOpen())
			{
				bool isHero = spec.Theme == IconTheme.HeroCoral || spec.Theme == IconTheme.HeroAmber || spec.Theme == IconTheme.HeroTerracotta;
				System.Windows.Media.Brush bgBrush;
				System.Windows.Media.Pen borderPen;
				System.Windows.Media.Brush fgBrush;

				switch (spec.Theme)
				{
				case IconTheme.HeroCoral:
					bgBrush = new LinearGradientBrush(
						System.Windows.Media.Color.FromRgb(255, 122, 69),
						System.Windows.Media.Color.FromRgb(194, 56, 18),
						new System.Windows.Point(0, 0),
						new System.Windows.Point(1, 1));
					borderPen = new System.Windows.Media.Pen(new SolidColorBrush(System.Windows.Media.Color.FromArgb(110, 255, 220, 190)), 1.5);
					fgBrush = System.Windows.Media.Brushes.White;
					break;
				case IconTheme.HeroAmber:
					bgBrush = new LinearGradientBrush(
						System.Windows.Media.Color.FromRgb(251, 191, 36),
						System.Windows.Media.Color.FromRgb(224, 79, 38),
						new System.Windows.Point(0, 0),
						new System.Windows.Point(1, 1));
					borderPen = new System.Windows.Media.Pen(new SolidColorBrush(System.Windows.Media.Color.FromArgb(110, 255, 240, 200)), 1.5);
					fgBrush = System.Windows.Media.Brushes.White;
					break;
				case IconTheme.HeroTerracotta:
					bgBrush = new LinearGradientBrush(
						System.Windows.Media.Color.FromRgb(234, 88, 12),
						System.Windows.Media.Color.FromRgb(154, 52, 18),
						new System.Windows.Point(0, 0),
						new System.Windows.Point(1, 1));
					borderPen = new System.Windows.Media.Pen(new SolidColorBrush(System.Windows.Media.Color.FromArgb(110, 255, 210, 180)), 1.5);
					fgBrush = System.Windows.Media.Brushes.White;
					break;
				case IconTheme.AmberDuotone:
					bgBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 247, 230));
					borderPen = new System.Windows.Media.Pen(new SolidColorBrush(System.Windows.Media.Color.FromRgb(253, 213, 150)), 2.0);
					fgBrush = new LinearGradientBrush(
						System.Windows.Media.Color.FromRgb(245, 158, 11),
						System.Windows.Media.Color.FromRgb(217, 70, 15),
						new System.Windows.Point(0, 0),
						new System.Windows.Point(1, 1));
					break;
				case IconTheme.CrimsonDuotone:
					bgBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 240, 236));
					borderPen = new System.Windows.Media.Pen(new SolidColorBrush(System.Windows.Media.Color.FromRgb(252, 195, 182)), 2.0);
					fgBrush = new LinearGradientBrush(
						System.Windows.Media.Color.FromRgb(239, 68, 68),
						System.Windows.Media.Color.FromRgb(185, 28, 28),
						new System.Windows.Point(0, 0),
						new System.Windows.Point(1, 1));
					break;
				case IconTheme.EmeraldDuotone:
					bgBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(238, 252, 244));
					borderPen = new System.Windows.Media.Pen(new SolidColorBrush(System.Windows.Media.Color.FromRgb(167, 243, 208)), 2.0);
					fgBrush = new LinearGradientBrush(
						System.Windows.Media.Color.FromRgb(16, 185, 129),
						System.Windows.Media.Color.FromRgb(21, 128, 61),
						new System.Windows.Point(0, 0),
						new System.Windows.Point(1, 1));
					break;
				default: // CoralDuotone
					bgBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 241, 230));
					borderPen = new System.Windows.Media.Pen(new SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 199, 171)), 2.0);
					fgBrush = new LinearGradientBrush(
						System.Windows.Media.Color.FromRgb(255, 107, 53),
						System.Windows.Media.Color.FromRgb(200, 62, 24),
						new System.Windows.Point(0, 0),
						new System.Windows.Point(1, 1));
					break;
				}

				// Draw rounded squircle card
				Rect cardRect = new Rect(2, 2, size - 4, size - 4);
				dc.DrawRoundedRectangle(bgBrush, borderPen, cardRect, 15, 15);

				// Subtle 3D glossy top highlight on Hero badges
				if (isHero)
				{
					LinearGradientBrush glossBrush = new LinearGradientBrush(
						System.Windows.Media.Color.FromArgb(75, 255, 255, 255),
						System.Windows.Media.Color.FromArgb(0, 255, 255, 255),
						new System.Windows.Point(0, 0),
						new System.Windows.Point(0, 1));
					dc.DrawRoundedRectangle(glossBrush, null, new Rect(4, 4, size - 8, (size - 8) * 0.48), 12, 12);
				}

				// Parse & fit Font Awesome geometry centered inside the badge
				Geometry geom = Geometry.Parse(spec.PathData);
				Rect bounds = geom.Bounds;
				if (!bounds.IsEmpty && bounds.Width > 0 && bounds.Height > 0)
				{
					double targetBox = isHero ? 36.0 : 38.0;
					double scale = Math.Min(targetBox / bounds.Width, targetBox / bounds.Height);
					double scaledW = bounds.Width * scale;
					double scaledH = bounds.Height * scale;
					double offsetX = (size - scaledW) / 2.0 - bounds.Left * scale;
					double offsetY = (size - scaledH) / 2.0 - bounds.Top * scale;

					TransformGroup tg = new TransformGroup();
					tg.Children.Add(new ScaleTransform(scale, scale));
					tg.Children.Add(new TranslateTransform(offsetX, offsetY));
					dc.PushTransform(tg);
					dc.DrawGeometry(fgBrush, null, geom);
					dc.Pop();
				}
			}

			RenderTargetBitmap rtb = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
			rtb.Render(visual);
			PngBitmapEncoder encoder = new PngBitmapEncoder();
			encoder.Frames.Add(BitmapFrame.Create(rtb));
			using (MemoryStream ms = new MemoryStream())
			{
				encoder.Save(ms);
				ms.Position = 0;
				using (Bitmap temp = new Bitmap(ms))
				{
					return new Bitmap(temp);
				}
			}
		}
		catch
		{
			return null;
		}
	}
}
