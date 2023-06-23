namespace Helpers

module Int =
    open System

    let fromNullable (nullable: Nullable<int>) =
        if nullable.HasValue then
            nullable.Value
        else
            0

