using System;

namespace Gryzak.Models
{
    public class ApiConfig
    {
        public string ApiUrl { get; set; } = "";
        public string ApiToken { get; set; } = "";
        public int ApiTimeout { get; set; } = 30;
        public string OrderListEndpoint { get; set; } = "/orders";
        public string OrderDetailsEndpoint { get; set; } = "/index.php?route=extension/module/orders&token=strefalicencji&order_id={order_id}&format=json";
    }
}

