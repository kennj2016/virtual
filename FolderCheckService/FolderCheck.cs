using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using SP4_API.Controllers;
using SP4_API.Entities;
using SP4_API.Models;
using System.IO;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Mvc;

namespace FolderCheckService
{
    public class FolderCheck : BackgroundService
    {
        private readonly ILogger<FolderCheck> _logger;
        private readonly SP4_API.Models.UNIT_TESTContext _context;

        public FolderCheck(ILogger<FolderCheck> logger)
        {
            _logger = logger;
        }
        public FolderCheck(SP4_API.Models.UNIT_TESTContext context)
        {
            _context = context;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("FolderCheck running at: {time}", DateTimeOffset.Now);
                // Folder check and fix db to match directories
                string rootPath = Path.Combine(Directory.GetCurrentDirectory(), "SusplanDocuments");

                List<string> result = new List<String>();

                try
                {
                    var Nodes = await _context.NodeTableCustom.FromSqlRaw($"SELECT * FROM VW_NODE_PATH").ToListAsync();

                    foreach (var item in Nodes)
                    {
                        string currPath = rootPath + item.ABS_PATH;

                        if (!System.IO.Directory.Exists(currPath))
                        {
                            var susplanNodes = await _context.SusplanNodes.FindAsync(item.NodeId);
                            _context.SusplanNodes.Remove(susplanNodes);
                            await _context.SaveChangesAsync();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    _logger.LogInformation("Error at: {time}", DateTimeOffset.Now, ex.Message);

                }
                
            }
            await Task.Delay(1000, stoppingToken);
        }
    }
}
