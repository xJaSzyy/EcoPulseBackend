namespace EcoPulseBackend.Enums;

public enum VehicleType
{
    /// <summary>
    /// Легковые автомобили
    /// </summary>
    Passenger = 1,
    
    /// <summary>
    /// Легковые дизельные
    /// </summary>
    DieselPassenger = 2,
    
    /// <summary>
    /// Грузовые карбюраторные с грузоподъемностью до 3 т (в том числе работающие на сжиженном нефтяном газе) и микроавтобусы
    /// </summary>
    CargoCarburetorLow = 3,
    
    /// <summary>
    /// Грузовые карбюраторные с грузоподъемностью более 3 т (в том числе работающие на сжиженном нефтяном газе)
    /// </summary>
    CargoCarburetorHigh = 4,
    
    /// <summary>
    /// Автобусы карбюраторные
    /// </summary>
    CarburetorBuses = 5,
    
    /// <summary>
    /// Грузовые дизельные
    /// </summary>
    DieselTrucks = 6,
    
    /// <summary>
    /// Автобусы дизельные
    /// </summary>
    DieselBuses = 7,
    
    /// <summary>
    /// Грузовые газобаллонные, работающие на сжатом природном
    /// </summary>
    CargoGas = 8,
}