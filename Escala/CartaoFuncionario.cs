namespace Escala
{
    /// <summary>
    /// Representa o cartão de itinerário de um funcionário para o dia
    /// </summary>
    public class CartaoFuncionario
    {
        /// <summary>
        /// Nome do funcionário
        /// </summary>
        public string? Nome { get; set; }
        
        /// <summary>
        /// Lista de itens do itinerário (horários e postos)
        /// </summary>
        public List<ItemItinerario> Itens { get; set; } = new List<ItemItinerario>();
    }
}
