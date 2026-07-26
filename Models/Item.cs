namespace Lab7.Models
{
    /// <summary>
    /// Материальная ценность кафедры (вариант 14)
    /// </summary>
    public class Item
    {
        public string Name { get; set; }           // Наименование предмета
        public int InventoryNumber { get; set; }   // Инвентарный номер (4 цифры)
        public int LabNumber { get; set; }         // Номер лаборатории
        public int PurchaseYear { get; set; }      // Год приобретения
        public int PurchaseMonth { get; set; }     // Месяц приобретения
        public decimal Cost { get; set; }          // Стоимость (грн.)
        public int ServiceLife { get; set; }       // Срок службы (лет)
        public int Quantity { get; set; }          // Количество

        // Год списания = год приобретения + срок службы
        public int WriteOffYear => PurchaseYear + ServiceLife;

        public string PurchaseDate => $"{PurchaseMonth:D2}/{PurchaseYear}";
    }
}
