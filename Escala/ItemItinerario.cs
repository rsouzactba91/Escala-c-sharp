namespace Escala
{
    /// <summary>
    /// Representa um item de horário e posto em um itinerário de funcionário
    /// </summary>
    public class ItemItinerario
    {
        /// <summary>
        /// Horário do item (ex: "08:00 x 08:40")
        /// </summary>
        public string? Horario { get; set; }
        
        /// <summary>
        /// Posto atribuído (ex: "VALET", "CAIXA", "QRF")
        /// </summary>
        public string? Posto { get; set; }
        
        /// <summary>
        /// Cor de fundo para exibição
        /// </summary>
        public System.Drawing.Color CorFundo { get; set; }
        
        /// <summary>
        /// Cor do texto para exibição
        /// </summary>
        public System.Drawing.Color CorTexto { get; set; }
    }
}
