// concepts.h - concepts
#include <concepts>
#include <type_traits>

#pragma once
template<class B1, class B2, class D>
concept either_base_of_v = (std::is_base_of_v<B1, D> || std::is_base_of_v<B2, D>);

// some_base_of_v<XLREF,XLREF12>
//template<class B1, class B2>
//concept any_base_of_v = either_base_of_v<B1, B2, D>;

template<class B, class D1, class D2>
concept both_derived_from_v = (std::is_base_of_v<B, D1> && std::is_base_of_v<B, D2>);

// concept all_derived_of_v
