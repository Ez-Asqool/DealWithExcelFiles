using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace DealWithExcelFiles.Migrations
{
    /// <inheritdoc />
    public partial class createProductsTable : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Products",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Band = table.Column<int>(type: "int", nullable: false),
                    CategoryCode = table.Column<string>(type: "nvarchar(8)", maxLength: 8, nullable: false),
                    Manufacturer = table.Column<string>(type: "nvarchar(16)", maxLength: 16, nullable: false),
                    PartSKU = table.Column<string>(type: "nvarchar(32)", maxLength: 32, nullable: false),
                    ItemDescription = table.Column<string>(type: "nvarchar(256)", maxLength: 256, nullable: false),
                    ListPrice = table.Column<string>(type: "nvarchar(16)", maxLength: 16, nullable: false),
                    MinimumDiscount = table.Column<string>(type: "nvarchar(8)", maxLength: 8, nullable: false),
                    DiscountedPrice = table.Column<string>(type: "nvarchar(16)", maxLength: 16, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Products", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Products");
        }
    }
}
